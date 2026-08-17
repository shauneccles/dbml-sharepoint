/**
 * dbml-sharepoint PROBE: CAN A DEPLOY SEE THAT A VIEW IS PROTECTED?
 *
 * #269. Writes a small list, three rows and three views, then reads three
 * SharePoint pages. Everything it creates is named after this run and is
 * removed by re-pasting with CLEANUP_LIST set.
 *
 * ANSWERED, over two runs on 2026-08-17. Yes: a view is protected when its
 * edit page returns HTTP 200 from the endpoint asked for, carries a sentinel,
 * and does not carry name="FieldPicker1". Re-run it when SharePoint changes
 * that page, because the check it supports is built on markup Microsoft
 * documents nowhere.
 *
 * THREE ROWS ARE STILL OPEN, added after run 2 in response to review, and one
 * of them matters more than the answered ones. S2, Q1 and C6 have never been
 * run. A third paste answers all three and costs one small list.
 *
 *   S2 IS THE ONE #267 RESTS ON. S1 cannot establish that the tautology is a
 *      tautology, and it was over-claiming when it said it could. Conjoined
 *      behind Eq(alpha) the left side already restricts to R1, so a group that
 *      wrongly excluded R2 or R3 returns the same single row and S1 still
 *      reports INERT. The partition is only visible when the group is asked ON
 *      ITS OWN. caml-chain-depth-probe.js T1 has exactly the same blind spot,
 *      for the same reason.
 *   Q1 reads the stored ViewQuery back. A view found by title is not proof it
 *      holds the tree that was sent, and SharePoint rewriting a filter on save
 *      is the behaviour this whole line of work is about.
 *   C6 asks a page that is not the editor whether it reads as protected. C5
 *      only ever showed a sentinel present on three pages that rendered, which
 *      is not the claim the guard depends on.
 *
 * Runs 1 and 2 remain valid for what they measured. S1 and F1 have since been
 * tightened as well: S1 now requires both queries to hit the control row
 * rather than merely agreeing with each other, and F1 now separates a redirect
 * and a transient failure from a refusal.
 *
 * WHY THIS EXISTS. #267 emits every view's filter in a shape the filter editor
 * refuses to open, because an editor that opens a filter of more than ten
 * conditions rewrites it to ten on save, silently and permanently. The shape
 * was measured on 2026-08-17 by caml-chain-depth-probe.js: a group in the
 * RIGHT child position is refused, and the tautology Or[IsNotNull(ID),
 * IsNull(ID)] both changes no rows (T1) and triggers the refusal (T2).
 *
 * The problem is that nothing can confirm it took. R1 measured
 * SP.View.ReadOnlyView reading false on an editable view and on a refused one
 * alike, and R2 measured a MERGE setting it true being accepted with HTTP 204
 * and read back false. No server-side property records the state. So the
 * deploy writes a protection it cannot verify, against this repository's rule
 * that anything which writes must read back and verify.
 *
 * G1 and G2 found the one surface that does record it. The view edit page
 * answers HTTP 200 to a same-origin fetch, and the refusal is served in the
 * HTML rather than painted by script, so a fetch can see it. What G2 could not
 * supply is a marker worth testing on, and its six candidates all failed.
 * Run 1 of this probe found one; the table below is why it had to be looked
 * for structurally rather than guessed at again:
 *
 *     complex filter      refused page only     English display text
 *     cannot be edited    refused page only     English display text
 *     ViewFilter          BOTH pages            does not discriminate
 *     FilterOnFieldName   NEITHER page
 *     onetidFilter        NEITHER page
 *     FilterOpt           NEITHER page
 *
 * THE DECISION THIS PROBE SERVES, taken 2026-08-17: ship English-only, and
 * where the marker is absent, warn that protection could not be verified
 * rather than failing the deploy. The view is still emitted protected; it is
 * the confirmation that is unavailable.
 *
 * THAT DECISION HAS A HOLE, AND F6 IS HOW IT GETS CLOSED. "Marker absent" is
 * also what a genuinely UNPROTECTED view looks like. If the emitted shape ever
 * stops triggering the refusal, through a SharePoint change or a regression in
 * how conditions are folded, every view reports "could not verify" and the
 * deploy stays quiet, with the real defect hidden inside the excuse. Reading
 * the tenant's own culture separates the two cases: not English means cannot
 * verify, English plus no marker means the protection did not take, and that
 * one is a failure. F6 asks whether the culture is readable at all.
 *
 * THE SHAPE UNDER TEST IS THE ONE #267 WILL EMIT, which is not the shape G2
 * measured. G2 compared a bare chain of twelve against And[IsNotNull(ID),
 * Or[chain12]]. What #267 emits for the 138 single-clause views is
 * And[leaf, tautology], a different tree, and a marker that appears for one
 * shape has not been shown to appear for the other. F2 re-asks it here.
 *
 * WHAT IS OBSERVED AND WHAT IS DEPENDED ON. F3 to F5 REPORT what differs
 * between the two pages. They assert no particular id, because the point is to
 * find one rather than to confirm a guess, and a probe that asserted its guess
 * would answer NOT ESTABLISHED the moment the guess was wrong and look
 * identical to the page being unreadable. P1 and P2 are the ground truth the
 * whole comparison rests on: an operator looks at both views and says which is
 * which. Without them a difference between two pages is just a difference.
 *
 * TWO CONTROLS, and neither is optional.
 *
 *   F8 asks whether a page is even deterministic. SharePoint pages carry
 *   request ids, tokens and timestamps. If two fetches of the SAME page differ,
 *   then a difference between two different pages proves nothing, and F3 to F5
 *   are noise being read as signal. It runs before them for that reason.
 *
 *   F7 fetches the edit page of a view with NO filter at all. An unfiltered
 *   view is editable, so whatever predicate comes out of this must classify it
 *   as unprotected. A predicate that calls it protected would report success
 *   for every view the tool never filtered.
 *
 * S1 is the third dependency. The two compared views must mean the same thing,
 * or their pages differ for a reason that has nothing to do with editability.
 * It re-measures T1's inertness on this fixture rather than citing it.
 *
 * RUN 1, 2026-08-17, revision 18f01ef7, one Team Site. Twelve questions,
 * TWELVE answered. P1 editable and P2 "complex filter", so the two pages
 * being compared really are an editable one and a refused one, and S1 INERT,
 * so they mean the same thing. The comparison is valid and what follows is
 * about the pages rather than about the filters.
 *
 * F5 IS THE RESULT. Thirty `name` attributes are on the editable page and
 * NONE are on the refused one:
 *
 *     FieldPicker1..FieldPicker10        10
 *     OperatorPicker1..OperatorPicker10  10
 *     NextIsAnd1..NextIsAnd9              9
 *     IsThereAQuery                       1
 *
 * These are ASP.NET form control names, not display text, so they do not move
 * with the tenant's language. The predicate is therefore ABSENCE: the filter
 * editor's controls are missing from a refused page, and that is what a deploy
 * can test for. C1 to C3 pin two of them and try to break the result.
 *
 * F4 CLOSES THE OTHER DIRECTION. Exactly one id is on the refused page and not
 * the editable one, `CssLink-cbec9c00...`, and F8 shows CssLink ids differing
 * between two fetches of the SAME page. It is noise. So there is no positive
 * marker unique to the refused page, and any check has to be built on what is
 * missing rather than on what is added.
 *
 * F8 BOUNDS THE NOISE rather than just reporting it. Two fetches of one page
 * differed by one character and by one CssLink-<guid> token each way, and
 * nothing else. So the id and name sets are otherwise reproducible, and the
 * only tokens that move are ones no predicate would use. C3 re-asks this of
 * the predicate itself, which is the version that matters.
 *
 * F7 IS THE CONTROL THAT MATTERS MOST. The unfiltered view's page carries the
 * same 139 ids the refused page lacks, so it reads as UNPROTECTED. Without
 * that, an absence test would have been detecting "this view has a filter"
 * and would have reported every plain list view as protected.
 *
 * F2: `complex filter` and `cannot be edited` discriminate on And[leaf,
 * tautology], the tree #267 emits, and not only on the chain G2 measured.
 * They stay unusable as a predicate on their own, being English.
 *
 * F6: `currentUICultureName` is "en-US", `currentLanguage` is 1033, and the
 * page's `<html lang>` is "en-AU". THE TWO DISAGREE, which is the useful part.
 * The UI culture is the user's and governs which language those display
 * strings come back in; the html lang follows the web's regional setting and
 * does not. A check that read the html attribute would be reading the wrong
 * one, and on this tenant it would have been wrong about the region while
 * still landing on English by luck.
 *
 * TEN SLOTS, AND WHERE THE CEILING COMES FROM. The editable page carries
 * exactly ten FieldPicker and ten OperatorPicker controls. U2 established the
 * ten-condition ceiling by an operator counting rows in the editor and
 * watching a save truncate a forty-condition filter to ten. The same ten is
 * sitting in the markup, so the ceiling is a property of the page rather than
 * of one observation, and C4 reads it out on any tenant without a person
 * looking.
 *
 * STILL OPEN: whether the absence predicate holds on a non-English tenant.
 * F6 establishes only that the culture value EXISTS. There is no non-English
 * tenant to hand, and the whole point of an absence test on a control NAME is
 * that it should not care, so this is a claim to confirm rather than one to
 * rely on.
 *
 * RUN 2, 2026-08-17, revision aa5eaae6, same site, after the candidate rows
 * were added. Seventeen questions, SEVENTEEN answered. The predicate is
 * settled, subject to the three rows added after this run:
 *
 *     A view is PROTECTED when its edit page comes back HTTP 200, carries a
 *     sentinel, and does NOT carry name="FieldPicker1".
 *
 * Every clause of that is measured and none of it reads display text.
 *
 *   C1 USABLE. FieldPicker1 is on the editable page, on the UNFILTERED page,
 *      and absent from the refused one. The unfiltered page is the half that
 *      matters: a control that merely tracked "this view has a filter" would
 *      satisfy a two-page comparison and then report every plain list view as
 *      protected.
 *   C2 AGREE. OperatorPicker1 gives the same verdict, so the result does not
 *      rest on one control name surviving a SharePoint update. They are
 *      independent controls, and disagreement between them is the signal that
 *      the markup has moved.
 *   C3 STABLE. All three pages re-fetched, same verdict. Run 1 measured a
 *      page differing from itself, so this is not a formality.
 *   C5 FOUND. `ViewFilter`, `ViewEdit` and `ctl00` are all on all three
 *      pages, so the check can require evidence the page ARRIVED before
 *      believing something is missing from it. Without that, a login
 *      redirect, an error page or a truncated response all lack FieldPicker1
 *      and read as protected, and the deploy reports success for a view it
 *      never saw.
 *
 * C1 IS ALSO EVIDENCE THE PAGE IS REFUSED, which is worth noticing because it
 * makes the run self-checking. If the guarded view had been editable its page
 * would carry FieldPicker1 and C1 would have answered NOT USABLE. So the two
 * halves cannot both be wrong in the same direction. P1 and P2 were confirmed
 * by an operator on run 1 against an identically constructed fixture; they
 * were not re-observed on run 2's list.
 *
 * C4: TEN SLOTS, and this is the second observation the ceiling wanted. The
 * editable page carries exactly ten FieldPicker and ten OperatorPicker
 * controls, matching the ten U2 established by an operator counting rows and
 * watching a save truncate forty conditions to that number. The editor has
 * ten condition slots, which is the mechanism rather than a coincidence.
 * Still ONE tenant: a second tenant, not a second run, is what would make ten
 * a constant worth asserting in limits.py.
 *
 * HOW TO RUN
 *   1. Open the target site as somebody who can create a list.
 *   2. Open a browser console on any page of that site.
 *   3. Set CONFIRMED = true and ALLOW_WRITES = true, then paste this file.
 *   4. Answer P1 and P2 by opening the two views named in their evidence and
 *      looking at the filter section of each view's settings.
 *   5. Copy the whole results block back verbatim.
 *   6. Re-paste with CLEANUP_LIST set to the list named at the end.
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
    //
    // MANUAL and NOT REACHED are counted open for the same reason. A MANUAL
    // row has SET UP an observation and is waiting for a person to make it,
    // so counting it answered lets a run print "0 not established" while
    // every browser check it asked for is still undone. That became visible
    // when the summary moved to the end of the run: before, those rows were
    // recorded after it and reported as NOT ESTABLISHED by accident.
    const OPEN_PREFIXES = ['NOT ESTABLISHED', 'SHORT', 'MANUAL', 'NOT REACHED'];
    const isOpen = (r) => OPEN_PREFIXES.some((p) => r.outcome.startsWith(p));
    const open = RESULTS.filter(isOpen).length;
    const waiting = RESULTS.filter(
      (r) => r.outcome.startsWith('MANUAL') || r.outcome.startsWith('NOT REACHED'),
    ).length;
    console.log(`${RESULTS.length} question(s); ${RESULTS.length - open} answered, ${open} open.`);
    if (waiting) {
      console.log(`${waiting} of those are waiting on an observation somebody has to make.`);
    }
    if (open) {
      console.log('A question with no observation is NOT a pass. Report it as open.');
    }
    console.log('Copy this whole block back verbatim.');
  };

  // Printed before any gate: a stale clipboard and a fix that did not
  // work produce identical transcripts otherwise.
  log('INFO', 'probe revision 8bd20c65. Quote this when reporting results.');

  // Set to the list named at the end of a previous run to drain and remove it.
  // The probe leaves its list behind so P1 and P2 can be looked at, and a
  // list with three views on it needs a way to be undone that is not deleting
  // things by hand from Site contents.
  const CLEANUP_LIST = '';

  // Run-unique so the probe never touches a list it did not create.
  const RUN = `${Date.now().toString(36)}`.slice(-6);
  const LIST_PREFIX = 'dbmlsp Probe ViewEdit ';
  const LIST = `${LIST_PREFIX}${RUN}`;
  const COL = 'Tag';

  // How many differing tokens to print per row. The pages are around half a
  // megabyte each and an uncapped symmetric difference would bury the result.
  // The COUNT is always reported, so a truncated list still says how much it
  // is showing of how many.
  const CAP = 40;

  const VERBOSE = { 'Content-Type': 'application/json;odata=verbose' };

  // Pinned from run 1, where F5 reported thirty `name` attributes on the
  // editable page and NONE on the refused one. Form control names, not
  // display text, so they do not move with the tenant's language.
  //
  // Two of them, because one is a single point of failure. If a SharePoint
  // update renames one control the other still answers, and C1 disagreeing
  // with C2 is itself the signal that the markup has moved.
  const CANDIDATES = ['FieldPicker1', 'OperatorPicker1'];

  // The predicate is ABSENCE, which fails dangerously on its own: a page
  // that did not render, an error page, or a truncated response all lack
  // the control and would read as protected. A sentinel present on EVERY
  // page, editable or not, is what separates "the editor is not here" from
  // "nothing is here". Reported rather than assumed, because a sentinel
  // that turns out to be conditional would silently disarm the guard.
  const SENTINELS = ['ViewFilter', 'ViewEdit', 'ctl00'];

  const eq = (v) => `<Eq><FieldRef Name="${COL}"/><Value Type="Text">${v}</Value></Eq>`;

  // The exact construct #267 emits. Every item has an ID, so the two halves
  // partition the rows and the conjunct changes nothing; it is a group, so on
  // the right it triggers the refusal. Both halves were measured by
  // caml-chain-depth-probe.js T1 and T2, and S1 re-measures the first here.
  const TAUTOLOGY = '<Or><IsNotNull><FieldRef Name="ID"/></IsNotNull>'
    + '<IsNull><FieldRef Name="ID"/></IsNull></Or>';
  const PLAIN_WHERE = eq('alpha');
  const GUARDED_WHERE = `<And>${eq('alpha')}${TAUTOLOGY}</And>`;

  const ROWS = [
    { Title: 'R1', [COL]: 'alpha' },
    { Title: 'R2', [COL]: 'beta' },
    { Title: 'R3', [COL]: 'gamma' },
  ];

  expect('Q0', 'the fixture actually built');
  expect('S1', 'does the guarded filter return the same rows as the plain one?');
  expect('S2', 'CONTROL: does the tautology ALONE return every row?');
  expect('Q1', 'did the guarded tree survive being STORED as a ViewQuery?');
  expect('F1', 'can the view edit page be fetched for both views?');
  expect('F8', 'CONTROL: is one page the same across two fetches?');
  expect('F2', 'do the English refusal markers appear on the shape #267 emits?');
  expect('F3', 'which id attributes are on the EDITABLE page and not the refused one?');
  expect('F4', 'which id attributes are on the REFUSED page and not the editable one?');
  expect('F5', 'which name attributes differ between the two pages?');
  expect('F6', 'can the tenant UI culture be read, so non-English can be told apart?');
  expect('F7', 'CONTROL: how does a view with NO filter compare?');
  expect('C1', 'is the pinned control name absent from the refused page and present on both others?');
  expect('C2', 'does the second candidate agree with the first?');
  expect('C3', 'is the predicate the same across a second fetch of all three pages?');
  expect('C4', 'how many condition slots does the editor markup carry?');
  expect('C5', 'is there a sentinel on every page, so a failed fetch cannot read as protected?');
  expect('C6', 'CONTROL: does a page that is NOT the editor read as protected?');
  expect('P1', 'GROUND TRUTH: does the plain view open its filter pane? (manual: look)');
  expect('P2', 'GROUND TRUTH: is the guarded view refused by the editor? (manual: look)');

  if (!CONFIRMED || !ALLOW_WRITES) {
    log('INFO', 'PLAN. Nothing has been touched.');
    log('INFO', `This probe would create the custom list '${LIST}' with a Text column`);
    log('INFO', `'${COL}', seed ${ROWS.length} rows, create three views (one unfiltered, one`);
    log('INFO', 'with a plain filter, one with that filter guarded by a tautology group),');
    log('INFO', 'then fetch each view\'s edit page and report what differs between them.');
    log('INFO', 'Set CONFIRMED = true and ALLOW_WRITES = true to run it.');
    report();
    return;
  }

  if (CLEANUP_LIST) {
    // Refuse a title this probe could not have created. The run-unique name
    // protects the RUN; it does nothing for a cleanup pasted later against a
    // mistyped or stale title, which would drain up to 5000 items from
    // somebody's list and recycle it.
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

  // ---- Fixture ----------------------------------------------------------
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
  const listPath = `web/lists/getbytitle('${LIST}')`;

  // A list that exists but has no usable column would answer every later
  // question with an empty view, so give up here rather than measuring one.
  const abandon = async (why) => {
    log('FAIL', `ABANDONING: ${why}.`);
    log('INFO', `Re-paste with CLEANUP_LIST = '${LIST}' to remove what was created.`);
    report();
  };

  // The `__metadata` type and the verbose Content-Type are both required.
  // caml-chain-depth-probe.js run 1 aborted with "The property 'Choices' does
  // not exist on type 'SP.Field'" for sending one without the other.
  digest = await getDigest();
  const field = await spPost(`${listPath}/fields`, {
    __metadata: { type: 'SP.FieldText' }, Title: COL, FieldTypeKind: 2, MaxLength: 64,
  }, digest, VERBOSE);
  if (!field.ok) {
    record('Q0', 'the fixture actually built', 'NOT ESTABLISHED',
      `the Text column could not be created: HTTP ${field.status} ${field.text.slice(0, 200)}.`);
    await abandon('the column could not be created');
    return;
  }

  const seedErrors = [];
  for (const row of ROWS) {
    digest = await getDigest();
    const wrote = await spPost(`${listPath}/items`, row, digest);
    if (!wrote.ok) seedErrors.push(`${row.Title}: HTTP ${wrote.status} ${wrote.text.slice(0, 120)}`);
  }

  // Read the fixture back rather than trusting the writes.
  const seeded = await spGet(`${listPath}/items?$select=Title,${COL}&$orderby=Id&$top=100`);
  const seenRows = (!readFailed(seeded) && seeded.body.value) || [];
  // Compare the PAIRS, not the titles. An item POST that succeeds while
  // dropping Tag leaves both filters returning nothing, and S1 would then
  // report INERT from two empty results, which is the premise failing rather
  // than the question being answered.
  const asPairs = (rows) => rows
    .map((r) => `${r.Title}=${r[COL] === null || r[COL] === undefined ? '<null>' : r[COL]}`)
    .sort();
  const seenPairs = asPairs(seenRows);
  const wantPairs = asPairs(ROWS);
  const fixtureOk = seedErrors.length === 0
    && JSON.stringify(seenPairs) === JSON.stringify(wantPairs);
  record('Q0', 'the fixture actually built',
    fixtureOk ? 'BUILT' : 'NOT ESTABLISHED',
    `rows=${seenRows.length}/${ROWS.length}; seed errors=${JSON.stringify(seedErrors)}; `
    + `read back ${JSON.stringify(seenPairs)} against ${JSON.stringify(wantPairs)}. Exactly one row `
    + "holds 'alpha', so both filtered views below must return exactly that row and any other count "
    + 'is the fixture, not the question.');
  if (!fixtureOk) {
    await abandon('the fixture did not build');
    return;
  }

  // ---- S1: the two views must MEAN the same thing -----------------------
  // Without this the pages differ for a reason that has nothing to do with
  // editability, and F3 to F5 attribute it to the wrong cause.
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

  const plainRows = await camlRows(PLAIN_WHERE);
  const guardedRows = await camlRows(GUARDED_WHERE);
  const bothQueried = plainRows.ok && guardedRows.ok;
  // Equality between the two is NOT enough. If GetItems ignored both filters,
  // or answered both with the same wrong rows, equality holds and the
  // tautology is crowned inert on the strength of a surface that is not
  // filtering at all. Both must also hit the control row the fixture
  // guarantees.
  const CONTROL = ['R1'];
  const hitsControl = bothQueried
    && JSON.stringify(plainRows.titles) === JSON.stringify(CONTROL)
    && JSON.stringify(guardedRows.titles) === JSON.stringify(CONTROL);
  const sameMeaning = hitsControl;
  record('S1', 'does the guarded filter return the same rows as the plain one?',
    !bothQueried ? 'NOT ESTABLISHED' : (sameMeaning ? 'INERT' : 'NOT INERT'),
    !bothQueried
      ? `plain: ${plainRows.error || 'ok'}; guarded: ${guardedRows.error || 'ok'}. Both are needed.`
      : `plain -> ${JSON.stringify(plainRows.titles)}, guarded -> ${JSON.stringify(guardedRows.titles)}, `
        + `control ${JSON.stringify(CONTROL)}. `
        + (sameMeaning
          ? 'The tautology changes nothing on this fixture, so the two views mean the same thing and any '
            + 'difference between their pages is about editability. This re-measures T1 rather than citing '
            + 'it, on the exact tree #267 emits.'
          : 'The two do not BOTH return the control row, so either the tautology changed the result or '
            + 'the query surface is not filtering. Either way every row below is comparing two different '
            + 'questions. Report this one first.'));


  // ---- S2: the tautology has to be a tautology --------------------------
  // S1 cannot see this. Conjoined behind Eq(alpha) the left side already
  // restricts to R1, so a tautology that wrongly excluded R2 or R3 would give
  // exactly the same answer and S1 would still say INERT. The partition is
  // only visible when the group is asked ON ITS OWN, where it must return
  // EVERY row.
  const tautRows = await camlRows(TAUTOLOGY);
  const allTitles = ROWS.map((r) => r.Title).sort();
  const partitions = tautRows.ok
    && JSON.stringify(tautRows.titles) === JSON.stringify(allTitles);
  record('S2', 'CONTROL: does the tautology ALONE return every row?',
    !tautRows.ok ? 'NOT ESTABLISHED' : (partitions ? 'PARTITIONS' : 'DOES NOT PARTITION'),
    !tautRows.ok
      ? `${tautRows.error}`
      : `${JSON.stringify(tautRows.titles)} against every row ${JSON.stringify(allTitles)}. `
        + (partitions
          ? 'IsNotNull(ID) and IsNull(ID) cover every row, so conjoining the group cannot remove one. This '
            + 'is the claim #267 rests on and S1 is structurally unable to make it.'
          : 'The group does NOT cover every row, so conjoining it would silently drop rows from any filter '
            + 'it is added to. #267 must not emit it. S1 can still read INERT here, which is exactly why '
            + 'this row exists.'));

  // ---- The three views --------------------------------------------------
  const makeView = async (title, where) => {
    const d = await getDigest();
    const payload = { Title: title, RowLimit: 100 };
    if (where) payload.ViewQuery = `<Where>${where}</Where>`;
    const v = await spPost(`${listPath}/views`, payload, d);
    return { ok: v.ok, status: v.status, text: v.text };
  };
  const viewErrors = [];
  for (const [title, where] of [
    ['Plain', PLAIN_WHERE], ['Guarded', GUARDED_WHERE], ['Unfiltered', null],
  ]) {
    const v = await makeView(title, where);
    if (!v.ok) viewErrors.push(`${title}: HTTP ${v.status} ${v.text.slice(0, 120)}`);
  }


  // ---- Q1: read the stored ViewQuery back --------------------------------
  // Finding a view by title says a view exists, not that it holds the tree
  // that was sent. SharePoint rewriting a filter on save is the behaviour this
  // whole line of work is about, so taking the POST's 200 as proof of the
  // stored shape would be assuming the thing under investigation.
  const stored = await spGet(`${listPath}/views?$select=Title,ViewQuery`);
  const storedBy = (title) => ((!readFailed(stored) && stored.body.value) || [])
    .find((v) => v.Title === title)?.ViewQuery ?? null;
  const norm = (s) => (s || '').replace(/\s+/g, '');
  const storedPlain = storedBy('Plain');
  const storedGuarded = storedBy('Guarded');
  const plainHeld = norm(storedPlain) === norm(`<Where>${PLAIN_WHERE}</Where>`);
  const guardedHeld = norm(storedGuarded) === norm(`<Where>${GUARDED_WHERE}</Where>`);
  const treesHeld = plainHeld && guardedHeld;
  record('Q1', 'did the guarded tree survive being STORED as a ViewQuery?',
    (storedPlain === null || storedGuarded === null) ? 'NOT ESTABLISHED'
      : (treesHeld ? 'SURVIVED' : 'REWRITTEN'),
    (storedPlain === null || storedGuarded === null)
      ? `Plain=${JSON.stringify(storedPlain)}, Guarded=${JSON.stringify(storedGuarded)}. One of the views `
        + 'could not be read back, so nothing below knows which tree it is looking at.'
      : `stored Plain: ${storedPlain}. stored Guarded: ${storedGuarded}. `
        + (treesHeld
          ? 'Both match what was sent, ignoring whitespace, so the pages below belong to the trees under '
            + 'test.'
          : 'SharePoint stored something OTHER than what was sent. Every page comparison below, and any '
            + 'manual look at P1 or P2, is then about a tree nobody chose. Report this before anything '
            + 'else.'));

  const listMeta = await spGet(`${listPath}?$select=Id`);
  const listGuid = (!readFailed(listMeta) && listMeta.body.Id) || null;
  const viewList = await spGet(`${listPath}/views?$select=Id,Title,ServerRelativeUrl`);
  const views = (!readFailed(viewList) && viewList.body.value) || [];
  const viewBy = (t) => views.find((v) => v.Title === t) || null;
  const plainView = viewBy('Plain');
  const guardedView = viewBy('Guarded');
  const unfilteredView = viewBy('Unfiltered');

  // ---- F1: is the page reachable at all? --------------------------------
  // A raw page fetch, not an _api call, so the harness helpers do not apply.
  // same-origin credentials so it carries the operator's session, which is
  // what a deploy running in the operator's browser would also have.
  const fetchEditPage = async (viewId) => {
    const url = `${WEB}/_layouts/15/ViewEdit.aspx?List=${encodeURIComponent(`{${listGuid}}`)}`
      + `&View=${encodeURIComponent(`{${viewId}}`)}`;
    try {
      const res = await fetch(url, { credentials: 'same-origin' });
      const text = await res.text();
      // A redirect to a login or to the modern settings surface answers 200,
      // so res.ok alone would hand the wrong HTML to every row below. The
      // page is only usable if the response came from the endpoint asked for.
      const landed = !res.redirected && res.url.includes('ViewEdit.aspx');
      return { ok: res.ok && landed, httpOk: res.ok, landed,
        status: res.status, length: text.length, text,
        redirected: res.redirected, finalUrl: res.url, error: null };
    } catch (err) {
      return { ok: false, httpOk: false, landed: false, status: 0, length: 0,
        text: '', redirected: false, finalUrl: null, error: String(err) };
    }
  };

  const haveIds = !!listGuid && !!plainView && !!guardedView;
  const pagePlain = haveIds ? await fetchEditPage(plainView.Id) : null;
  const pageGuarded = haveIds ? await fetchEditPage(guardedView.Id) : null;
  const bothFetched = !!pagePlain?.ok && !!pageGuarded?.ok;
  // A throttle, an auth failure or a network error is not the endpoint
  // saying no, and recording it as REFUSED would preserve a transient as
  // evidence that this approach is closed. isRefusal draws that line for the
  // REST helpers already; the same line applies to a page fetch.
  const pageOutcome = () => {
    if (bothFetched) return 'FETCHED';
    for (const page of [pagePlain, pageGuarded]) {
      if (page && page.httpOk && !page.landed) return 'REDIRECTED';
    }
    const refused = [pagePlain, pageGuarded].some((p) => p && isRefusal(p.status));
    return refused ? 'REFUSED' : 'NOT ESTABLISHED (transient)';
  };
  record('F1', 'can the view edit page be fetched for both views?',
    !haveIds ? 'NOT ESTABLISHED' : pageOutcome(),
    !haveIds
      ? `list id=${listGuid}, Plain=${plainView?.Id || null}, Guarded=${guardedView?.Id || null}. `
        + `View creation errors: ${JSON.stringify(viewErrors)}.`
      : `Plain: HTTP ${pagePlain.status}, ${pagePlain.length} chars`
        + `${pagePlain.redirected ? `, REDIRECTED to ${pagePlain.finalUrl}` : ''}`
        + `${pagePlain.error ? `, threw ${pagePlain.error}` : ''}. `
        + `Guarded: HTTP ${pageGuarded.status}, ${pageGuarded.length} chars`
        + `${pageGuarded.redirected ? `, REDIRECTED to ${pageGuarded.finalUrl}` : ''}`
        + `${pageGuarded.error ? `, threw ${pageGuarded.error}` : ''}. `
        + 'A REDIRECTED outcome means the classic page is not what an authenticated fetch gets and the '
        + 'approach closes here rather than at F3. A transient failure is held OPEN rather than recorded '
        + 'as a refusal, because a throttle is not the endpoint saying no.');

  // ---- F8: is a page even deterministic? --------------------------------
  // Runs before the diffs it validates. Two fetches of the SAME page, so any
  // difference is markup that varies per request rather than per view.
  const pagePlainAgain = bothFetched ? await fetchEditPage(plainView.Id) : null;
  const attrValues = (text, attr) => {
    const out = new Set();
    for (const m of text.matchAll(new RegExp(`\\b${attr}="([^"]{1,120})"`, 'g'))) out.add(m[1]);
    return out;
  };
  // Any token carrying one of these differs between the two pages by
  // construction, because they are different views on the same list. Left in,
  // they would fill the symmetric difference with noise that means nothing.
  const guidBits = [listGuid, plainView?.Id, guardedView?.Id, unfilteredView?.Id]
    .filter(Boolean)
    .flatMap((g) => [String(g).toLowerCase(), String(g).replace(/-/g, '').toLowerCase()]);
  const isNoise = (tok) => guidBits.some((g) => tok.toLowerCase().includes(g));
  const minus = (a, b) => [...a].filter((x) => !b.has(x) && !isNoise(x)).sort();
  const show = (list) => `${list.length} token(s)`
    + (list.length ? `, showing ${Math.min(list.length, CAP)}: ${JSON.stringify(list.slice(0, CAP))}` : '');

  const replayOk = !!pagePlainAgain?.ok;
  const idsPlain = bothFetched ? attrValues(pagePlain.text, 'id') : new Set();
  const idsPlain2 = replayOk ? attrValues(pagePlainAgain.text, 'id') : new Set();
  const driftA = replayOk ? minus(idsPlain, idsPlain2) : [];
  const driftB = replayOk ? minus(idsPlain2, idsPlain) : [];
  const stable = replayOk && driftA.length === 0 && driftB.length === 0;
  record('F8', 'CONTROL: is one page the same across two fetches?',
    !replayOk ? 'NOT ESTABLISHED' : (stable ? 'STABLE' : 'VARIES'),
    !replayOk
      ? 'the second fetch of the plain view\'s page did not succeed, so nothing below has a control.'
      : `lengths ${pagePlain.length} then ${pagePlainAgain.length}. ids only in the first: ${show(driftA)}. `
        + `ids only in the second: ${show(driftB)}. `
        + (stable
          ? 'The id set is reproducible, so a difference between two DIFFERENT pages is attributable to '
            + 'the pages rather than to the request.'
          : 'The id set VARIES between fetches of the same page, so those ids cannot carry a predicate and '
            + 'F3 and F4 must be read as including this much noise. Any candidate marker has to be checked '
            + 'against this list first.'));

  // ---- F2: the English markers, on the shape #267 emits ------------------
  // G2 measured And[IsNotNull(ID), Or[chain12]]. #267 emits And[leaf,
  // tautology] for the 138 single-clause views, which is a different tree.
  const MARKERS = ['complex filter', 'cannot be edited', 'CannotEditFilter', 'ViewFilter'];
  const markersOn = (page) => (page && page.ok
    ? MARKERS.filter((m) => page.text.includes(m))
    : null);
  const mPlain = markersOn(pagePlain);
  const mGuarded = markersOn(pageGuarded);
  const discriminating = bothFetched
    ? MARKERS.filter((m) => mGuarded.includes(m) && !mPlain.includes(m))
    : [];
  record('F2', 'do the English refusal markers appear on the shape #267 emits?',
    !bothFetched ? 'NOT ESTABLISHED' : (discriminating.length ? 'PRESENT' : 'ABSENT'),
    !bothFetched
      ? 'both pages are needed to say whether a marker discriminates.'
      : `editable page carries ${JSON.stringify(mPlain)}; guarded page carries ${JSON.stringify(mGuarded)}; `
        + `discriminating: ${JSON.stringify(discriminating)}. `
        + (discriminating.length
          ? 'So the refusal is visible for the tree #267 actually emits, not only for the chain G2 measured. '
            + 'These are still English display strings and cannot be the whole predicate; F6 is what makes '
            + 'them safe to use.'
          : 'NOTHING discriminates. Either this tree is not refused, which contradicts T2 and P2 will say '
            + 'so, or the refusal is expressed differently for it. Read P2 before concluding either.'));

  // ---- F3, F4, F5: what actually differs --------------------------------
  // Reported, never asserted. The point is to find a candidate, and a probe
  // that asserted one would answer NOT ESTABLISHED whenever the guess was
  // wrong, which is indistinguishable from the page being unreadable.
  const idsGuarded = bothFetched ? attrValues(pageGuarded.text, 'id') : new Set();
  const onlyEditable = bothFetched ? minus(idsPlain, idsGuarded) : [];
  const onlyRefused = bothFetched ? minus(idsGuarded, idsPlain) : [];
  record('F3', 'which id attributes are on the EDITABLE page and not the refused one?',
    !bothFetched ? 'NOT ESTABLISHED' : (onlyEditable.length ? 'FOUND' : 'NONE'),
    !bothFetched
      ? 'both pages are needed.'
      : `${show(onlyEditable)}. These are the most promising direction: the refused page is the SMALLER of `
        + `the two (${pageGuarded.length} against ${pagePlain.length} chars), so what is missing from it is `
        + 'the filter editor itself. A predicate reading "this control is ABSENT, therefore protected" is '
        + 'language-independent in a way the refusal sentence is not. Pick one, then confirm it over a '
        + 'second run before anything relies on it.');
  record('F4', 'which id attributes are on the REFUSED page and not the editable one?',
    !bothFetched ? 'NOT ESTABLISHED' : (onlyRefused.length ? 'FOUND' : 'NONE'),
    !bothFetched
      ? 'both pages are needed.'
      : `${show(onlyRefused)}. A container holding the refusal message would show up here, and its id would `
        + 'be a positive test that is not English text.');

  const namesPlain = bothFetched ? attrValues(pagePlain.text, 'name') : new Set();
  const namesGuarded = bothFetched ? attrValues(pageGuarded.text, 'name') : new Set();
  record('F5', 'which name attributes differ between the two pages?',
    !bothFetched ? 'NOT ESTABLISHED' : 'REPORTED',
    !bothFetched
      ? 'both pages are needed.'
      : `only on editable: ${show(minus(namesPlain, namesGuarded))}. `
        + `only on refused: ${show(minus(namesGuarded, namesPlain))}. `
        + 'Form control names survive markup changes better than generated ids do, so a candidate here is '
        + 'worth more than one from F3.');

  // ---- F6: can non-English be told apart from a failure? ----------------
  // The decision is to ship English-only and warn when the marker is absent.
  // That warning is only honest if "not English" can be distinguished from
  // "the protection did not take", which look identical at the marker.
  const ctxCulture = pageCtx.currentUICultureName || null;
  const ctxLanguage = pageCtx.currentLanguage || null;
  const langAttr = bothFetched
    ? (pagePlain.text.match(/<html[^>]*\blang="([^"]{2,12})"/i) || [])[1] || null
    : null;
  // The html lang follows the WEB's regional setting; the UI culture is the
  // user's and is what decides which language those display strings arrive
  // in. Run 1 measured them disagreeing, en-AU against en-US, so accepting
  // the attribute as a fallback would authorise the language branch on the
  // wrong surface. It is reported, never relied on.
  const cultureReadable = !!ctxCulture || !!ctxLanguage;
  record('F6', 'can the tenant UI culture be read, so non-English can be told apart?',
    !cultureReadable ? 'NOT ESTABLISHED' : 'READABLE',
    `_spPageContextInfo.currentUICultureName=${JSON.stringify(ctxCulture)}, `
    + `currentLanguage=${JSON.stringify(ctxLanguage)}. OBSERVED ONLY, not part of the verdict: `
    + `<html lang>=${JSON.stringify(langAttr)}. `
    + (cultureReadable
      ? 'So a deploy can tell an unverifiable tenant from a failed protection: a non-English culture warns, '
        + 'and English with no marker is a FAILURE rather than a warning. Both readings need a second '
        + 'observation on a non-English tenant before the branch is trusted, and there is no such tenant '
        + 'here, so what this row establishes is only that the value EXISTS.'
      : 'No UI-culture surface is readable, so "marker absent" cannot be attributed and every unprotected '
        + 'view would be excused as a language difference. The warning would then hide the defect it was '
        + 'meant to report. An <html lang> present here does not rescue it: that is the web\'s region, '
        + 'not the language the strings come back in.'));

  // ---- F7: the negative control -----------------------------------------
  // An unfiltered view is editable. Any predicate that calls it protected
  // would report success for every view this tool never filtered.
  const pageUnfiltered = (haveIds && unfilteredView)
    ? await fetchEditPage(unfilteredView.Id) : null;
  const unfOk = !!pageUnfiltered?.ok;
  record('F7', 'CONTROL: how does a view with NO filter compare?',
    !unfOk ? 'NOT ESTABLISHED' : 'REPORTED',
    !unfOk
      ? `the unfiltered view's page was not fetched (view=${unfilteredView?.Id || null}).`
      : `${pageUnfiltered.length} chars against ${pagePlain.length} editable and ${pageGuarded.length} `
        + `refused. markers: ${JSON.stringify(markersOn(pageUnfiltered))}. `
        + `ids it has that the refused page does not: ${show(minus(attrValues(pageUnfiltered.text, 'id'), idsGuarded))}. `
        + 'It must read as UNPROTECTED under whatever predicate is chosen. If it looks like the refused '
        + 'page, the predicate is detecting something other than the refusal.');

  // ---- C1..C5: pin the candidate and try to break it ---------------------
  // F3 to F5 are discovery and report everything. These five take the one
  // candidate that survived and ask whether it behaves like a predicate.
  const namesUnfiltered = unfOk ? attrValues(pageUnfiltered.text, 'name') : null;
  const threePages = bothFetched && unfOk;

  // The predicate under test, stated as the three-way it has to satisfy: the
  // control is on the editable page, on the unfiltered page, and NOT on the
  // refused one. Two-way agreement is not enough. A control absent from the
  // unfiltered page too would mean the test detects "has a filter" rather than
  // "is protected", and would report every plain list view as protected.
  const verdictFor = (name) => {
    if (!threePages) return null;
    const onEditable = namesPlain.has(name);
    const onRefused = namesGuarded.has(name);
    const onUnfiltered = namesUnfiltered.has(name);
    return { name, onEditable, onRefused, onUnfiltered,
      usable: onEditable && onUnfiltered && !onRefused };
  };
  const v1 = verdictFor(CANDIDATES[0]);
  const v2 = verdictFor(CANDIDATES[1]);
  const describe = (v) => (v
    ? `${v.name}: editable=${v.onEditable}, unfiltered=${v.onUnfiltered}, refused=${v.onRefused}`
    : 'not measured');

  record('C1', 'is the pinned control name absent from the refused page and present on both others?',
    !threePages ? 'NOT ESTABLISHED' : (v1.usable ? 'USABLE' : 'NOT USABLE'),
    !threePages
      ? 'all three pages are needed: the predicate is a three-way, not a comparison of two.'
      : `${describe(v1)}. `
        + (v1.usable
          ? 'So `name="' + v1.name + '" is absent` distinguishes a protected view from both an editable '
            + 'filtered view and an unfiltered one, without reading any display text. Pair it with C5\'s '
            + 'sentinel before it is used, because absence alone cannot tell a protected page from a page '
            + 'that never arrived.'
          : 'It does not satisfy the three-way, so it is not the predicate. C2 may still be.'));

  record('C2', 'does the second candidate agree with the first?',
    !threePages ? 'NOT ESTABLISHED' : ((v1.usable === v2.usable) ? 'AGREE' : 'DISAGREE'),
    !threePages
      ? 'all three pages are needed.'
      : `${describe(v2)}. `
        + ((v1.usable === v2.usable)
          ? 'Both candidates give the same verdict, so the result does not rest on one control name '
            + 'surviving a SharePoint update.'
          : 'The two candidates DISAGREE, which means the markup has moved or one of them is conditional. '
            + 'Neither should be used until this is understood.'));

  // Run 1 measured the id set varying between two fetches of the same page,
  // by exactly one CssLink-<guid> token each way. That is noise, but it means
  // "stable" cannot be assumed for the tokens a predicate would read.
  const refetchNames = async (view) => {
    const page = await fetchEditPage(view.Id);
    return page.ok ? attrValues(page.text, 'name') : null;
  };
  const again = threePages
    ? {
      plain: await refetchNames(plainView),
      guarded: await refetchNames(guardedView),
      unfiltered: await refetchNames(unfilteredView),
    }
    : null;
  const stableFor = (name) => (again && again.plain && again.guarded && again.unfiltered
    ? again.plain.has(name) && again.unfiltered.has(name) && !again.guarded.has(name)
    : null);
  const s1 = stableFor(CANDIDATES[0]);
  const s2 = stableFor(CANDIDATES[1]);
  const replayed = s1 !== null && s2 !== null;
  record('C3', 'is the predicate the same across a second fetch of all three pages?',
    !replayed ? 'NOT ESTABLISHED' : ((s1 === (v1 && v1.usable) && s2 === (v2 && v2.usable)) ? 'STABLE' : 'VARIES'),
    !replayed
      ? 'one of the three pages did not come back on the second fetch, so there is no replay to compare.'
      : `second fetch: ${CANDIDATES[0]} usable=${s1}, ${CANDIDATES[1]} usable=${s2}; `
        + `first fetch: ${v1.usable}, ${v2.usable}. `
        + ((s1 === v1.usable && s2 === v2.usable)
          ? 'Two observations agree, which is the minimum before anything relies on this.'
          : 'The verdict CHANGED between two fetches of the same three pages, so it is not a predicate at '
            + 'all and no amount of tenant-side care would make it one.'));

  // Independent of the predicate, and worth its own row. U2 established the
  // ten-condition ceiling by an operator counting rows in the editor. The
  // editor's own markup should carry that number, which makes it measurable on
  // any tenant instead of by eye.
  const slotCount = (names, prefix) => [...names].filter((n) => new RegExp(`^${prefix}\\d+$`).test(n)).length;
  const fieldSlots = bothFetched ? slotCount(namesPlain, 'FieldPicker') : null;
  const opSlots = bothFetched ? slotCount(namesPlain, 'OperatorPicker') : null;
  record('C4', 'how many condition slots does the editor markup carry?',
    !bothFetched ? 'NOT ESTABLISHED' : `${fieldSlots} SLOTS`,
    !bothFetched
      ? 'the editable page is needed.'
      : `FieldPicker1..n = ${fieldSlots}, OperatorPicker1..n = ${opSlots}. `
        + 'U2 established the ceiling by an operator counting conditions in the editor and watching a save '
        + 'truncate to that number. This is the same number read out of the markup, so it can be measured '
        + 'on any tenant without a person looking. If it matches, the ceiling is a property of the page '
        + 'rather than of one observation.');

  const onAll = threePages
    ? SENTINELS.filter((s) => pagePlain.text.includes(s)
      && pageGuarded.text.includes(s) && pageUnfiltered.text.includes(s))
    : [];
  record('C5', 'is there a sentinel on every page, so a failed fetch cannot read as protected?',
    !threePages ? 'NOT ESTABLISHED' : (onAll.length ? 'FOUND' : 'NONE'),
    !threePages
      ? 'all three pages are needed.'
      : `present on all three: ${JSON.stringify(onAll)}; of candidates ${JSON.stringify(SENTINELS)}. `
        + (onAll.length
          ? 'So a check can require the sentinel BEFORE believing the control is absent. What this row '
            + 'shows is only that the sentinel is PRESENT on three pages that rendered; that it is absent '
            + 'from a page that did not is a separate claim, and C6 is where it gets tested.'
          : 'NOTHING is on all three pages, so absence of the control cannot be distinguished from absence '
            + 'of the page. The predicate must not be built until a sentinel is found.'));


  // ---- C6: the way an absence test fails --------------------------------
  // C5 shows a sentinel on three pages that all rendered. It cannot show the
  // sentinel is ABSENT from a page that did not, which is the claim the guard
  // actually depends on. So ask for a view that does not exist and report what
  // comes back: if that response carries a sentinel and no FieldPicker1, the
  // guard classifies a non-page as a protected view.
  const BOGUS_VIEW = '00000000-0000-0000-0000-000000000001';
  const bogus = haveIds ? await fetchEditPage(BOGUS_VIEW) : null;
  const bogusSentinels = bogus ? SENTINELS.filter((s) => bogus.text.includes(s)) : [];
  const bogusNames = bogus && bogus.text ? attrValues(bogus.text, 'name') : new Set();
  const bogusHasControl = bogusNames.has(CANDIDATES[0]);
  // The guard as it would be written, run against a page that is not the one
  // asked for. It must NOT come out protected.
  const bogusReadsProtected = !!bogus && bogus.ok
    && bogusSentinels.length > 0 && !bogusHasControl;
  record('C6', 'CONTROL: does a page that is NOT the editor read as protected?',
    !bogus ? 'NOT ESTABLISHED' : (bogusReadsProtected ? 'FALSE POSITIVE' : 'REJECTED'),
    !bogus
      ? 'the list ids were not available, so no bogus request was made.'
      : `HTTP ${bogus.status}, landed=${bogus.landed}, ${bogus.length} chars; sentinels present `
        + `${JSON.stringify(bogusSentinels)}; ${CANDIDATES[0]} present=${bogusHasControl}. `
        + (bogusReadsProtected
          ? 'The guard would call this protected. It is not a view at all. The sentinel is too weak, or the '
            + 'landed check is not being applied, and the check must not ship in this form.'
          : 'The guard rejects it, so a request that does not reach the editor cannot be mistaken for a '
            + 'protected view. This is one shape of failure, not all of them: a response truncated after '
            + 'the sentinel and before the controls is still unmeasured, and a length or completeness test '
            + 'is what would close that.'));

  // ---- P1, P2: the ground truth ------------------------------------------
  record('P1', 'GROUND TRUTH: does the plain view open its filter pane? (manual: look)',
    plainView ? 'MANUAL' : 'NOT ESTABLISHED',
    plainView
      ? `OPEN ${window.location.origin}${plainView.ServerRelativeUrl}, then its view settings, and report `
        + 'ONE of: "filter pane" (editable) or "complex filter" (refused). This is expected to be editable, '
        + 'and if it is not then the comparison has no editable side and every row above is measuring two '
        + 'refused pages.'
      : 'the Plain view was not created, so there is nothing to open.');
  record('P2', 'GROUND TRUTH: is the guarded view refused by the editor? (manual: look)',
    guardedView ? 'MANUAL' : 'NOT ESTABLISHED',
    guardedView
      ? `OPEN ${window.location.origin}${guardedView.ServerRelativeUrl}, then its view settings, and report `
        + 'ONE of: "filter pane" or "complex filter". THIS IS THE ROW EVERYTHING ELSE RESTS ON. If it opens '
        + 'the filter pane then the tautology does not protect this tree, F2 to F5 are comparing two '
        + 'editable pages, and #267 must not emit this shape.'
      : 'the Guarded view was not created, so there is nothing to open.');

  report();
  log('INFO', `KEEPING '${LIST}' so P1 and P2 can be looked at.`);
  log('INFO', `When finished, re-paste this file with CLEANUP_LIST = '${LIST}' to drain and remove it.`);
})();
