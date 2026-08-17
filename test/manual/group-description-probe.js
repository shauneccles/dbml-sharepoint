/**
 * dbml-sharepoint PROBE: WHAT A SITE GROUP'S SETTINGS DO ON THE WAY BACK
 *
 * WHY THIS EXISTS. `templates/deploy/_security_principals.js.j2` writes five
 * fields onto every site group it creates or adopts (`Description` and four
 * membership flags) and NEVER READS ONE OF THEM BACK. The create path POSTs
 * and logs "created". The reconcile path checks `mergeResp.ok` and logs
 * "declared membership controls reconciled". An HTTP 200 says the request was
 * accepted; it does not say the tenant stored what was sent.
 *
 * That is the one thing AGENTS.md refuses outright: anything that writes must
 * read back and verify. Lists already do. A list Description is composed,
 * written, read back and byte-compared, and the characters that did not
 * survive that comparison are refused at build time. Groups got none of it,
 * because nobody had measured the surface.
 *
 * `SP.Group.Description` is a DIFFERENT SURFACE from `SP.List.Description`.
 * The list findings are not evidence for it and must not be reused as such.
 * This file measures the group surface on its own.
 *
 * WHAT TURNS ON THE ANSWER
 *
 *   - Whether the deploy can gain a read-back verification at all. If the
 *     description does not round-trip, "verify" has to mean something other
 *     than byte-equality, and the deploy needs to know which.
 *   - Issue #211, a provenance marker on groups. Lists carry "Provisioned by
 *     dbml-sharepoint from family/entity." and `assess.js` reports one that
 *     has gone. Groups carry nothing, so nothing can tell a group this tool
 *     created from one that was already on the site.
 *   - Issue #209, the Critical that #211 would unblock. The deploy adopts any
 *     existing group it finds BY NAME and grants it Full Control, with no
 *     membership gate. The gate everyone reaches for ("adopt silently if we
 *     made it, fail closed if we did not") cannot be written until a group
 *     can carry a mark that survives being written and read.
 *
 * SEPARATING WHAT IS DEPENDED ON FROM WHAT IS OBSERVED. G6 writes a long
 * description and REPORTS the length that comes back; it does not assert a
 * ceiling. A control that asserted over the value it exists to discover would
 * fail the moment the measurement started working, and that looks identical
 * to a real refusal. The same applies to G4 and G5: a character that does not
 * survive is an ANSWER, recorded FAIL, not a broken run.
 *
 * THE NEGATIVE CONTROL IS N1 AND IT IS NOT OPTIONAL. Every other question
 * here reads a group back and compares. If a read of a group that does not
 * exist came back 200, every one of those comparisons would be meaningless
 * and the probe would certify the surface on the strength of a bug. N1 asks
 * for the absent group first.
 *
 * WHAT THIS TOUCHES. One site group, created and deleted by this probe under
 * a RUN-UNIQUE name it prints before doing anything. It touches no list, no
 * item, and no group it did not create, and if the name is somehow already
 * taken it stops without writing or deleting. It needs ManagePermissions, so
 * a site owner.
 *
 * RUN 1 (2026-08-13), one Microsoft 365 group-connected Team Site. Ten
 * questions, ten answered.
 *
 *   ANSWERED AND SETTLED:
 *     G1, G2  a group Description round-trips BYTE-IDENTICAL, through both
 *             the create path and the MERGE path.
 *     G3      MERGE is PARTIAL: a field left out of the request is
 *             preserved. Confirmed non-vacuously: the flag the request did
 *             carry flipped in the same call.
 *     G4, G5  an ampersand and a run of two spaces BOTH survive. The
 *             restrictions a LIST note carries do NOT transfer to a group.
 *     G6      512 characters, and the server REFUSES rather than truncating:
 *             HTTP 500, "The parameter Description cannot be null or bigger
 *             than 512 characters." Now enforced at build time by
 *             `group_description_too_long`.
 *     G7      the four membership flags round-trip at CREATE.
 *     G8      a marker of the shape lists use survives, so #211 is possible
 *             and #209's adoption gate has something to key on.
 *
 *   AMBIGUOUS AFTER RUN 1, AND THE PROBE'S FAULT: G9 reported
 *   AutoAcceptRequestToJoinLeave sent true, read false, with the other
 *   three taking. It sent AllowRequestToJoinLeave FALSE in the same
 *   request (a contradictory pair, since a group cannot auto-accept join
 *   requests it does not accept), so it could not distinguish a lost write
 *   from a coerced dependent setting. G10 and G11 were added to settle
 *   that and the empty-description question.
 *
 * RUN 2 (2026-08-14), same site. Twelve questions, twelve answered. Every
 * run-1 result reproduced identically, and both open questions closed:
 *
 *     G10     with AllowRequestToJoinLeave TRUE, AutoAccept TOOK. So run 1
 *             saw SharePoint COERCING a dependent setting, not the
 *             reconcile path losing a write. The server accepts the
 *             contradictory pair with HTTP 200 and then stores auto-accept
 *             as false. That silent correction is now refused at build
 *             time by `group_auto_accept_without_requests`, and G9 no
 *             longer compares that flag. It would report FAIL for correct
 *             behaviour on every future run.
 *     G11     an EMPTY description is accepted and reads back "". The
 *             server's "cannot be null" does not extend to the empty
 *             string, so a mapping that omits `description:` (which the
 *             loader turns into "") is known-good rather than unproven.
 *
 * NOTHING IS OPEN. The file still runs end to end; re-confirming the
 * settled questions costs nothing and would catch a tenant-specific
 * answer, which two runs against ONE site cannot rule out.
 *
 * HOW TO RUN
 *   1. Open the target site as a SITE OWNER.
 *   2. Open a browser console on any page of that site.
 *   3. Set CONFIRMED = true and ALLOW_WRITES = true, then paste this file.
 *   4. Copy the whole results block back verbatim.
 *
 * WHEN FINISHED: the probe deletes the group it created. If it aborted early,
 * delete the group whose name it printed in its first line, from Site
 * settings > Site permissions. The name carries a run-unique suffix, so it
 * will not collide with anything the site already had.
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
  log('INFO', 'probe revision 780552ec. Quote this when reporting results.');

  // RUN-UNIQUE, and that is a safety property rather than tidiness.
  //
  // An earlier draft used a fixed name and deleted it before starting, to
  // guarantee the negative control had nothing to find. On a site where that
  // name already existed (an interrupted run, or an administrator who
  // happened to choose it), the probe would have destroyed a real group and
  // its grants before looking at it, while its own preamble promised it
  // touches no group it did not create. Review caught it; this is the fix.
  //
  // The suffix makes collision effectively impossible, and the existence
  // check below still refuses to proceed rather than assuming.
  const RUN = `${Date.now().toString(36)}`.slice(-6);
  const GROUP = `dbmlsp Probe Description ${RUN}`;

  // Mirrors `analysis/list_description.py`'s MARKER_TEMPLATE shape, because
  // G8 asks whether THAT string could live on a group. A marker that cannot
  // survive the round trip cannot be the basis of #209's adoption gate.
  const MARKER = 'Provisioned by dbml-sharepoint from probe/Group.';

  const PLAIN = 'dbmlsp probe plain description';
  const MERGED = 'dbmlsp probe merged description';
  const AMPERSAND = 'dbmlsp probe risks & issues';
  const SPACES = 'dbmlsp probe two  spaces';
  const LONG = `dbmlsp probe long ${'x'.repeat(1000)}`;

  expect('N1', 'CONTROL: does reading a group that does not exist actually fail?');
  expect('G1', 'Does a plain Description written at CREATE read back byte-identical?');
  expect('G2', 'Does a Description written by MERGE read back byte-identical?');
  expect('G3', 'Does a MERGE that OMITS Description preserve the previous one?');
  expect('G4', 'Does an ampersand survive the round trip?');
  expect('G5', 'Does a run of two spaces survive the round trip?');
  expect('G6', 'What length comes back when 1000+ characters go out?');
  expect('G7', 'Do the four membership flags read back as written at CREATE?');
  expect('G8', 'Does a provenance marker of the shape lists use survive?');
  expect('G9', 'Do the three independent membership flags read back as written by MERGE?');
  expect('G10', 'Is AutoAccept refused only when its prerequisite is off, or always?');
  expect('G11', 'Is an EMPTY description accepted?');

  if (!CONFIRMED || !ALLOW_WRITES) {
    log('INFO', 'PLAN: nothing has been touched.');
    log('INFO', `This probe would create the site group '${GROUP}', write and`);
    log('INFO', 'read back its Description several ways, then delete it.');
    log('INFO', 'Set CONFIRMED = true and ALLOW_WRITES = true to run it.');
    report();
    return;
  }

  // The group settings surface is the one place this project has PROVEN
  // works: _security_principals.js.j2 creates groups against live tenants
  // with a verbose body. Mirroring it exactly keeps a failure here meaning
  // "the description did not survive" rather than "the probe framed the
  // request differently from the deployer".
  const VERBOSE = { 'Content-Type': 'application/json;odata=verbose' };

  const readGroup = async () =>
    spGet(`web/sitegroups/getbyname('${encodeURIComponent(GROUP)}')`);

  const deleteGroup = async (digest) =>
    spPost(
      `web/sitegroups/removebyloginname(@v)?@v='${encodeURIComponent(GROUP)}'`,
      {}, digest,
    );

  let digest = await getDigest();
  log('INFO', `This run uses the group name '${GROUP}'.`);

  // ---- N1: the negative control, and the existence check, in one read ---
  // Nothing is deleted first. If the name is somehow taken, that is a group
  // this probe did not create and it stops rather than touching it.
  const absent = await readGroup();
  if (absent.ok) {
    record('N1', 'CONTROL: does reading a group that does not exist actually fail?',
      'NOT ESTABLISHED (name already taken)',
      `'${GROUP}' already exists on this site. Nothing was written or deleted. Re-run to get a different name, or delete that group by hand if a previous run left it.`);
    report();
    return;
  }
  // PASS on 404 ONLY. 401/403 are about WHO is asking and 408/429 about the
  // moment (the harness's own isRefusal contract says so), and any of them
  // would otherwise certify the prerequisite that gives every comparison
  // below its meaning. A throttled request must not stand in for absence.
  if (absent.status !== 404) {
    record('N1', 'CONTROL: does reading a group that does not exist actually fail?',
      `NOT ESTABLISHED (HTTP ${absent.status})`,
      'the control needs a clean 404 for a name that is not there; this status does not establish absence, so nothing below would mean anything');
    report();
    return;
  }
  record('N1', 'CONTROL: does reading a group that does not exist actually fail?',
    'PASS',
    'HTTP 404 for a group that is not there, so a successful read below means something');

  // ---- G1: create with a plain description ------------------------------
  digest = await getDigest();
  const created = await spPost('web/sitegroups', {
    __metadata: { type: 'SP.Group' },
    Title: GROUP,
    Description: PLAIN,
    AllowMembersEditMembership: false,
    AllowRequestToJoinLeave: false,
    AutoAcceptRequestToJoinLeave: false,
    OnlyAllowMembersViewMembership: true,
  }, digest, VERBOSE);
  if (!created.ok) {
    record('BOOT', 'create the probe group',
      `NOT ESTABLISHED (HTTP ${created.status})`, created.text.slice(0, 300));
    report();
    return;
  }

  const afterCreate = await readGroup();
  if (readFailed(afterCreate)) {
    record('G1', 'Does a plain Description written at CREATE read back byte-identical?',
      `NOT ESTABLISHED (HTTP ${afterCreate.status})`, 'the group was created but could not be read back');
  } else {
    const got = afterCreate.body.Description;
    record('G1', 'Does a plain Description written at CREATE read back byte-identical?',
      got === PLAIN ? 'PASS' : 'FAIL',
      `sent ${JSON.stringify(PLAIN)}, read ${JSON.stringify(got)}`);
  }

  // ---- G7: the four flags the deploy writes and never verifies ----------
  if (!readFailed(afterCreate)) {
    const g = afterCreate.body;
    const wanted = {
      AllowMembersEditMembership: false,
      AllowRequestToJoinLeave: false,
      AutoAcceptRequestToJoinLeave: false,
      OnlyAllowMembersViewMembership: true,
    };
    const wrong = Object.keys(wanted).filter((k) => g[k] !== wanted[k]);
    record('G7', 'Do the four membership flags read back as written at CREATE?',
      wrong.length === 0 ? 'PASS' : 'FAIL',
      wrong.length === 0
        ? 'all four match what was sent'
        : `differ: ${wrong.map((k) => `${k} sent ${wanted[k]}, read ${g[k]}`).join('; ')}`);
  }

  // A MERGE, read back, reported against what was sent. Every remaining
  // question is this shape, so it is one function rather than five copies.
  const mergeAndRead = async (id, question, body, expected, describe) => {
    const d = await getDigest();
    const res = await spPost(
      `web/sitegroups/getbyname('${encodeURIComponent(GROUP)}')`,
      { __metadata: { type: 'SP.Group' }, ...body },
      d,
      { ...VERBOSE, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
    );
    if (!res.ok) {
      record(id, question,
        isRefusal(res.status) ? 'FAIL' : `NOT ESTABLISHED (HTTP ${res.status})`,
        `the MERGE itself came back HTTP ${res.status}: ${res.text.slice(0, 200)}`);
      return null;
    }
    const back = await readGroup();
    if (readFailed(back)) {
      record(id, question, `NOT ESTABLISHED (HTTP ${back.status})`,
        'the MERGE succeeded but the group could not be read back');
      return null;
    }
    const got = back.body.Description;
    record(id, question, describe ? describe(got) : (got === expected ? 'PASS' : 'FAIL'),
      `sent ${JSON.stringify(expected)}, read ${JSON.stringify(got)}`);
    return got;
  };

  // ---- G2: a description written by the reconcile path ------------------
  await mergeAndRead('G2', 'Does a Description written by MERGE read back byte-identical?',
    { Description: MERGED }, MERGED);

  // ---- G3: is MERGE partial, or does omitting a field clear it? ---------
  // The shared-group design in #210 depends on this: two families reconcile
  // the same object, so whether an omitted field is preserved or wiped
  // decides what a second deploy does to the first one's settings.
  // The omitted-field MERGE must CHANGE something, or this question passes
  // vacuously: a request that did nothing at all would also leave the
  // Description intact, and "preserved" would mean nothing. `create` set
  // AllowRequestToJoinLeave false, so flipping it true is a change the
  // read-back can confirm actually landed.
  const d3 = await getDigest();
  const partial = await spPost(
    `web/sitegroups/getbyname('${encodeURIComponent(GROUP)}')`,
    { __metadata: { type: 'SP.Group' }, AllowRequestToJoinLeave: true },
    d3,
    { ...VERBOSE, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
  );
  if (!partial.ok) {
    record('G3', 'Does a MERGE that OMITS Description preserve the previous one?',
      `NOT ESTABLISHED (HTTP ${partial.status})`, partial.text.slice(0, 200));
  } else {
    const back = await readGroup();
    if (readFailed(back)) {
      record('G3', 'Does a MERGE that OMITS Description preserve the previous one?',
        `NOT ESTABLISHED (HTTP ${back.status})`, 'could not read the group back');
    } else {
      const got = back.body.Description;
      const landed = back.body.AllowRequestToJoinLeave === true;
      record('G3', 'Does a MERGE that OMITS Description preserve the previous one?',
        !landed
          ? 'NOT ESTABLISHED (the MERGE changed nothing)'
          : (got === MERGED ? 'PASS' : 'FAIL'),
        !landed
          ? 'AllowRequestToJoinLeave did not flip, so this request did not take effect and "preserved" would prove nothing'
          : (got === MERGED
            ? 'the flag flipped AND the omitted Description survived, so MERGE is partial'
            : `the flag flipped but the omitted Description became ${JSON.stringify(got)}, so an omitted field is NOT preserved`));
    }
  }

  // ---- G4/G5: characters the LIST surface refuses -----------------------
  await mergeAndRead('G4', 'Does an ampersand survive the round trip?',
    { Description: AMPERSAND }, AMPERSAND);
  await mergeAndRead('G5', 'Does a run of two spaces survive the round trip?',
    { Description: SPACES }, SPACES);

  // ---- G6: length, OBSERVED rather than asserted ------------------------
  await mergeAndRead('G6', 'What length comes back when 1000+ characters go out?',
    { Description: LONG }, LONG,
    (got) => (got === LONG
      ? 'PASS'
      : (typeof got === 'string' ? `OBSERVED (came back ${got.length} of ${LONG.length})` : 'FAIL')));

  // ---- G8: the marker shape #211 would use ------------------------------
  await mergeAndRead('G8', 'Does a provenance marker of the shape lists use survive?',
    { Description: MARKER }, MARKER);

  // ---- G9: the flags through the RECONCILE path -------------------------
  // G7 only proves the CREATE path. Phase 1.2's reconcile is the one that
  // matters and the one nothing verifies: its whole purpose is that "a
  // pre-existing group with the right name but permissive flags must not be
  // accepted as compliant". Sending a value that is already correct would
  // pass without establishing anything, which is what G7 alone did.
  //
  // So every one of the four is the OPPOSITE of the state the run has
  // reached by here: create set (false, false, false, true) and G3 flipped
  // AllowRequestToJoinLeave to true, leaving (false, true, false, true).
  // A MERGE that silently ignored any of them reads back the old value.
  const flipped = {
    AllowMembersEditMembership: true,
    AllowRequestToJoinLeave: false,
    AutoAcceptRequestToJoinLeave: true,
    OnlyAllowMembersViewMembership: false,
  };
  const d9 = await getDigest();
  const flipRes = await spPost(
    `web/sitegroups/getbyname('${encodeURIComponent(GROUP)}')`,
    { __metadata: { type: 'SP.Group' }, ...flipped },
    d9,
    { ...VERBOSE, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
  );
  if (!flipRes.ok) {
    record('G9', 'Do the three independent membership flags read back as written by MERGE?',
      isRefusal(flipRes.status) ? 'FAIL' : `NOT ESTABLISHED (HTTP ${flipRes.status})`,
      `the MERGE came back HTTP ${flipRes.status}: ${flipRes.text.slice(0, 200)}`);
  } else {
    const back9 = await readGroup();
    if (readFailed(back9)) {
      record('G9', 'Do the three independent membership flags read back as written by MERGE?',
        `NOT ESTABLISHED (HTTP ${back9.status})`, 'the MERGE succeeded but the group could not be read back');
    } else {
      // AutoAccept is EXCLUDED from the comparison, and that is a finding
      // rather than an exemption. This request sends AllowRequestToJoinLeave
      // false, and G10 established across two runs that SharePoint then
      // stores AutoAccept as false whatever was asked for -- a group cannot
      // auto-accept requests it does not accept. Comparing it here would
      // report FAIL for correct behaviour on every future run, which is how
      // a probe teaches its reader to ignore it. G10 owns that flag; this
      // question owns the three that are independent.
      const independent = { ...flipped };
      delete independent.AutoAcceptRequestToJoinLeave;
      const wrong9 = Object.keys(independent).filter((k) => back9.body[k] !== independent[k]);
      const coerced = back9.body.AutoAcceptRequestToJoinLeave === false;
      record('G9', 'Do the three independent membership flags read back as written by MERGE?',
        wrong9.length === 0 ? 'PASS' : 'FAIL',
        wrong9.length === 0
          ? `all three took, so the reconcile path persists flags${coerced ? '; AutoAccept read false as G10 predicts, its prerequisite being off' : '; NOTE AutoAccept did NOT coerce this time, which contradicts G10. Report it'}`
          : `did NOT take: ${wrong9.map((k) => `${k} sent ${independent[k]}, read ${back9.body[k]}`).join('; ')}`);
    }
  }

  // ---- G10: disambiguate what G9 saw ------------------------------------
  // RUN 1 (2026-08-13) had G9 report AutoAcceptRequestToJoinLeave sent true,
  // read false, with the other three taking. That has two explanations and
  // they call for opposite responses:
  //
  //   (a) the reconcile path silently loses that flag -- a real defect, and
  //       the deploy could not trust any flag it writes; or
  //   (b) SharePoint COERCES it, because G9 sent AllowRequestToJoinLeave
  //       false in the same request and a group cannot auto-accept join
  //       requests it does not accept. A dependent setting being corrected
  //       is not a lost write.
  //
  // G9 cannot tell them apart, because it sent the contradictory pair. That
  // was a flaw in the probe, not a finding about the tenant. G10 sends the
  // COHERENT pair -- both true -- so a false read-back here means (a) and a
  // true read-back means (b).
  const coherent = {
    AllowRequestToJoinLeave: true,
    AutoAcceptRequestToJoinLeave: true,
  };
  const d10 = await getDigest();
  const cohRes = await spPost(
    `web/sitegroups/getbyname('${encodeURIComponent(GROUP)}')`,
    { __metadata: { type: 'SP.Group' }, ...coherent },
    d10,
    { ...VERBOSE, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
  );
  if (!cohRes.ok) {
    record('G10', 'Is AutoAccept refused only when its prerequisite is off, or always?',
      isRefusal(cohRes.status) ? 'FAIL' : `NOT ESTABLISHED (HTTP ${cohRes.status})`,
      `the MERGE came back HTTP ${cohRes.status}: ${cohRes.text.slice(0, 200)}`);
  } else {
    const back10 = await readGroup();
    if (readFailed(back10)) {
      record('G10', 'Is AutoAccept refused only when its prerequisite is off, or always?',
        `NOT ESTABLISHED (HTTP ${back10.status})`, 'could not read the group back');
    } else {
      const took = back10.body.AutoAcceptRequestToJoinLeave === true;
      record('G10', 'Is AutoAccept refused only when its prerequisite is off, or always?',
        took ? 'PASS' : 'FAIL',
        took
          ? 'with AllowRequestToJoinLeave TRUE the flag took, so G9 saw a dependent setting being coerced, not a lost write'
          : 'the flag did not take even with its prerequisite on, so the reconcile path really does lose it. The deploy cannot trust a flag it writes');
    }
  }

  // ---- G11: is an empty description accepted? ---------------------------
  // The server named null in the same breath as the 512 ceiling, and
  // `mapping_loader` turns an omitted `description:` into "". Nobody has
  // measured whether that empty string is treated as null and refused, so a
  // mapping that simply omits the key is currently unproven rather than
  // known-good.
  const d11 = await getDigest();
  const emptyRes = await spPost(
    `web/sitegroups/getbyname('${encodeURIComponent(GROUP)}')`,
    { __metadata: { type: 'SP.Group' }, Description: '' },
    d11,
    { ...VERBOSE, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
  );
  if (!emptyRes.ok) {
    record('G11', 'Is an EMPTY description accepted?',
      isRefusal(emptyRes.status) ? 'FAIL' : `NOT ESTABLISHED (HTTP ${emptyRes.status})`,
      `HTTP ${emptyRes.status}: ${emptyRes.text.slice(0, 200)}`);
  } else {
    const back11 = await readGroup();
    record('G11', 'Is an EMPTY description accepted?',
      readFailed(back11) ? `NOT ESTABLISHED (HTTP ${back11.status})` : 'PASS',
      readFailed(back11)
        ? 'the MERGE succeeded but the group could not be read back'
        : `accepted; read back ${JSON.stringify(back11.body.Description)}`);
  }

  // ---- Clean up ---------------------------------------------------------
  digest = await getDigest();
  const removed = await deleteGroup(digest);
  log(removed.ok ? 'OK' : 'FAIL',
    removed.ok
      ? `Deleted '${GROUP}'.`
      : `Could not delete '${GROUP}' (HTTP ${removed.status}). Remove it by hand.`);

  report();
  console.log('');
  console.log('=== WHAT TO DO WITH THIS ===');
  console.log('G1/G2 decide whether the deploy can gain a byte-compare read-back');
  console.log('of a group Description (AGENTS.md requires one; it has none today).');
  console.log('G3 decides what a second family deploying to the same site does to');
  console.log("the first one's group settings (see #210).");
  console.log('G4/G5 say whether the list-note restrictions apply to groups too.');
  console.log('G8 decides whether #211, and the #209 adoption gate that depends');
  console.log('on it, are possible at all.');
})();
