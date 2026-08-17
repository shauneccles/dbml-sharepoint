/**
 * dbml-sharepoint PROBE: does a list Description survive a write and read?
 *
 * WHY. `reconcileListDescription` in `templates/deploy/_field_reconcile.js.j2`
 * writes the composed Description, reads it back and compares the two byte for
 * byte. Whether SharePoint returns a list Description unchanged was inferred
 * from the field case rather than measured. `entity_note_may_not_round_trip`
 * refused an ampersand, a line break and a run of spaces in every table note
 * because of that gap, and both the rule and the deploy template named this
 * probe as the thing that would settle it.
 *
 * SCOPE. `SP.List.Description` only. `test/manual/group-description-probe.js`
 * measured `SP.Group.Description`, which is a different property on a
 * different type: a group refuses over 512 characters and a list accepted
 * 1018, so neither result is evidence for the other.
 *
 * WHAT IT TOUCHES. One custom list, created and deleted under a run-unique
 * name the probe prints before doing anything. It touches no group, no
 * permission and no list it did not create. If the name is already taken it
 * stops without writing or deleting.
 *
 * L7 REPORTS A LOWER BOUND, not the absence of a limit. It sends one length
 * and does not search for a ceiling.
 *
 * RUN 1, 2026-08-14, revision 25d4b223, one Microsoft 365 group-connected
 * Team Site. Ten questions, ten answered, all PASS.
 *
 *     L1, L2   byte-identical through the create path and the MERGE path
 *     L3       a MERGE omitting Description preserves it, so MERGE is
 *              partial. The Title the request did carry changed in the same
 *              call, so the result is not vacuous.
 *     L4       an ampersand survives, and does not return as `&amp;`
 *     L5       a run of two spaces survives, uncollapsed
 *     L6       a bare LF survives, unrewritten
 *     L7       1018 characters returned intact, so the accepted length is at
 *              least 1018
 *     L8       an empty description is accepted
 *     L9       the composed note-plus-marker shape survives
 *
 * RUN 2, 2026-08-14, revision 2759ce32, same site. Eleven questions, eleven
 * answered. Every run-1 result reproduced, and L10 closed the last one: a
 * CRLF survives byte-identical, with neither half stripped.
 *
 * OUTCOME. All four characters `entity_note_may_not_round_trip` refused
 * survive, so the rule was deleted along with the caveat in
 * `website/docs/reference/dbml.md`. Issue #203 asked the same question about
 * the marker's own components and is answered by the same runs.
 *
 * L5 sent ASCII spaces only. A whitespace run containing a tab or a
 * non-breaking space is still refused, by
 * `entity_note_whitespace_unmeasured`, because nothing has measured those.
 *
 * HOW TO RUN
 *   1. Open the target site as somebody who can create a list.
 *   2. Open a browser console on any page of that site.
 *   3. Set CONFIRMED = true and ALLOW_WRITES = true, then paste this file.
 *   4. Copy the whole results block back verbatim.
 *
 * WHEN FINISHED: the probe deletes the list it created. If it aborted early,
 * delete the list whose name it printed in its first line.
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

  // Identifies which version was pasted, since a stale clipboard and a failed fix read the same.
  log('INFO', 'probe revision 1f136997. Quote this when reporting results.');

  // Run-unique so the probe never touches a list it did not create.
  const RUN = `${Date.now().toString(36)}`.slice(-6);
  const LIST = `dbmlsp Probe Description ${RUN}`;

  const PLAIN = 'dbmlsp probe plain description';
  const MERGED = 'dbmlsp probe merged description';
  // Asked separately so a partial answer narrows the rule instead of leaving it whole.
  const AMPERSAND = 'dbmlsp probe risks & issues';
  const SPACES = 'dbmlsp probe two  spaces';
  // L5 measured two. The guard allows any length on the reasoning that
  // collapsing is one behaviour, and this observes it rather than reasons.
  const LONG_SPACES = `dbmlsp probe eight${' '.repeat(8)}spaces`;
  const NEWLINE = 'dbmlsp probe first line\nsecond line';
  // Built from char codes so nothing between here and the console normalises the bytes.
  const CR = String.fromCharCode(13);
  const LF = String.fromCharCode(10);
  const CRLF_TEXT = `dbmlsp probe first line${CR}${LF}second line`;
  const LONG = `dbmlsp probe long ${'x'.repeat(1000)}`;
  // The shape the tool writes: a note, then MARKER_TEMPLATE from list_description.py.
  const COMPOSED = 'Risks and issues for the service. Provisioned by dbml-sharepoint from probe/Risk.';

  expect('N1', 'CONTROL: does reading a list that does not exist actually fail?');
  expect('L1', 'Does a plain Description read back byte-identical from CREATE?');
  expect('L2', 'Does a Description written by MERGE read back byte-identical?');
  expect('L3', 'Does a MERGE that OMITS Description preserve the previous one?');
  expect('L4', 'Does an ampersand survive, or come back as &amp;?');
  expect('L5', 'Does a run of two spaces survive, or get collapsed?');
  expect('L6', 'Does a line break survive, or get rewritten?');
  expect('L7', 'What comes back when 1000+ characters go out?');
  expect('L8', 'Is an EMPTY description accepted?');
  expect('L9', 'Does the composed note-plus-marker shape survive intact?');
  expect('L10', 'Does a CRLF survive, given a bare LF does?');
  expect('L11', 'Does a LONGER run of spaces survive, given two do?');

  if (!CONFIRMED || !ALLOW_WRITES) {
    log('INFO', 'PLAN. Nothing has been touched.');
    log('INFO', `This probe would create the custom list '${LIST}', write and`);
    log('INFO', 'read back its Description several ways, then delete it.');
    log('INFO', 'Set CONFIRMED = true and ALLOW_WRITES = true to run it.');
    report();
    return;
  }

  // Mirrors the deployer's request shape, so a failure means the value did not survive.
  const VERBOSE = { 'Content-Type': 'application/json;odata=verbose' };

  const readList = async () => spGet(`web/lists/getbytitle('${encodeURIComponent(LIST)}')`);

  const mergeDescription = async (value) => {
    const d = await getDigest();
    return spPost(
      `web/lists/getbytitle('${encodeURIComponent(LIST)}')`,
      { __metadata: { type: 'SP.List' }, Description: value },
      d,
      { ...VERBOSE, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
    );
  };

  log('INFO', `This run uses the list name '${LIST}'.`);

  // ---- N1: negative control and existence check in one read -------------
  const absent = await readList();
  if (absent.ok) {
    record('N1', 'CONTROL: does reading a list that does not exist actually fail?',
      'NOT ESTABLISHED (name already taken)',
      `'${LIST}' already exists. Nothing was written or deleted. Re-run for a different name.`);
    report();
    return;
  }
  if (absent.status !== 404) {
    record('N1', 'CONTROL: does reading a list that does not exist actually fail?',
      `NOT ESTABLISHED (HTTP ${absent.status})`,
      'the control needs a clean 404; 401/403 are about who is asking and 408/429 about the moment, and none of them establish absence');
    report();
    return;
  }
  record('N1', 'CONTROL: does reading a list that does not exist actually fail?',
    'PASS', 'HTTP 404 for a list that is not there, so a successful read below means something');

  // ---- L1: create ------------------------------------------------------
  const dc = await getDigest();
  const created = await spPost('web/lists', {
    __metadata: { type: 'SP.List' },
    Title: LIST,
    BaseTemplate: 100,
    Description: PLAIN,
  }, dc, VERBOSE);
  if (!created.ok) {
    record('BOOT', 'create the probe list',
      `NOT ESTABLISHED (HTTP ${created.status})`, created.text.slice(0, 300));
    report();
    return;
  }
  const afterCreate = await readList();
  if (readFailed(afterCreate)) {
    record('L1', 'Does a plain Description read back byte-identical from CREATE?',
      `NOT ESTABLISHED (HTTP ${afterCreate.status})`, 'created but could not be read back');
  } else {
    const got = afterCreate.body.Description;
    record('L1', 'Does a plain Description read back byte-identical from CREATE?',
      got === PLAIN ? 'PASS' : 'FAIL',
      `sent ${JSON.stringify(PLAIN)}, read ${JSON.stringify(got)}`);
  }

  // One MERGE, one read, reported against what was sent.
  const mergeAndRead = async (id, question, value, describe) => {
    const res = await mergeDescription(value);
    if (!res.ok) {
      record(id, question,
        isRefusal(res.status) ? 'FAIL' : `NOT ESTABLISHED (HTTP ${res.status})`,
        `the MERGE came back HTTP ${res.status}: ${res.text.slice(0, 220)}`);
      return null;
    }
    const back = await readList();
    if (readFailed(back)) {
      record(id, question, `NOT ESTABLISHED (HTTP ${back.status})`,
        'the MERGE succeeded but the list could not be read back');
      return null;
    }
    const got = back.body.Description;
    record(id, question, describe ? describe(got) : (got === value ? 'PASS' : 'FAIL'),
      `sent ${JSON.stringify(value)}, read ${JSON.stringify(got)}`);
    return got;
  };

  const l2Read = await mergeAndRead(
    'L2', 'Does a Description written by MERGE read back byte-identical?', MERGED);

  // ---- L3: partial MERGE ------------------------------------------------
  // L3 asks whether an omitted Description is preserved, and compares against
  // MERGED, so it is only meaningful if L2 actually left MERGED in place. A
  // throttled L2 leaves PLAIN there, and L3 would then report FAIL for a
  // setup that never happened.
  if (l2Read !== MERGED) {
    record('L3', 'Does a MERGE that OMITS Description preserve the previous one?',
      'NOT ESTABLISHED (L2 did not set up the value)',
      `this question compares against ${JSON.stringify(MERGED)}, which L2 was supposed to leave in place. L2 read ${JSON.stringify(l2Read)} instead, so a comparison here would report on the setup rather than on SharePoint.`);
  } else {
    // Changes the Title as well, or a request that did nothing would also look like preservation.
    const RENAMED = `${LIST} b`;
    const d3 = await getDigest();
    const partial = await spPost(
      `web/lists/getbytitle('${encodeURIComponent(LIST)}')`,
      { __metadata: { type: 'SP.List' }, Title: RENAMED },
      d3,
      { ...VERBOSE, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
    );
    if (!partial.ok) {
      record('L3', 'Does a MERGE that OMITS Description preserve the previous one?',
        `NOT ESTABLISHED (HTTP ${partial.status})`, partial.text.slice(0, 220));
    } else {
      // getbytitle follows the NEW title once the rename lands.
      const back = await spGet(`web/lists/getbytitle('${encodeURIComponent(RENAMED)}')`);
      if (readFailed(back)) {
        record('L3', 'Does a MERGE that OMITS Description preserve the previous one?',
          `NOT ESTABLISHED (HTTP ${back.status})`, 'could not read the list back after the rename');
      } else {
        const landed = back.body.Title === RENAMED;
        const got = back.body.Description;
        record('L3', 'Does a MERGE that OMITS Description preserve the previous one?',
          !landed ? 'NOT ESTABLISHED (the MERGE changed nothing)'
            : (got === MERGED ? 'PASS' : 'FAIL'),
          !landed
            ? 'the Title did not change, so this request did not take effect and "preserved" would prove nothing'
            : (got === MERGED
              ? 'the Title changed AND the omitted Description survived, so MERGE is partial'
              : `the Title changed but the Description became ${JSON.stringify(got)}, so an omitted field is NOT preserved`));
      }

      // Restored and verified unconditionally, because a list left renamed
      // makes L4 to L10 report FAIL for a 404 they caused themselves.
      // Reading LIST back is the right check either way: if the rename never
      // landed, the restoring MERGE 404s harmlessly and the read still passes.
      const dr = await getDigest();
      await spPost(
        `web/lists/getbytitle('${encodeURIComponent(RENAMED)}')`,
        { __metadata: { type: 'SP.List' }, Title: LIST },
        dr,
        { ...VERBOSE, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
      );
      const restored = await readList();
      if (readFailed(restored)) {
        record('BOOT', 'restore the probe list\'s title after L3',
          `NOT ESTABLISHED (HTTP ${restored.status})`,
          `the list did not answer to '${LIST}' again after L3's rename, so every question below would address a title that may not exist and report a plumbing failure as a SharePoint verdict. Stopping instead. Look for a list named '${RENAMED}' and delete it by hand.`);
        report();
        return;
      }
    }
  }

  // ---- L4/L5/L6: the three the build refuses today ----------------------
  // Each reports the returned value, so a transformation is visible rather than only a failure.
  await mergeAndRead('L4', 'Does an ampersand survive, or come back as &amp;?', AMPERSAND,
    (got) => (got === AMPERSAND ? 'PASS'
      : (typeof got === 'string' && got.includes('&amp;') ? 'FAIL (escaped to &amp;)' : 'FAIL')));
  await mergeAndRead('L5', 'Does a run of two spaces survive, or get collapsed?', SPACES,
    (got) => (got === SPACES ? 'PASS'
      : (typeof got === 'string' && got.includes(' two spaces') ? 'FAIL (run collapsed)' : 'FAIL')));
  await mergeAndRead('L6', 'Does a line break survive, or get rewritten?', NEWLINE,
    (got) => (got === NEWLINE ? 'PASS'
      : (typeof got === 'string' && got.includes('\r\n') ? 'FAIL (rewritten to CRLF)' : 'FAIL')));

  // ---- L7: length, OBSERVED ---------------------------------------------
  // Says "lower bound" in the outcome, because a bare PASS here reads as "no limit".
  await mergeAndRead('L7', 'What comes back when 1000+ characters go out?', LONG,
    (got) => (got === LONG
      ? `PASS (lower bound only: ${LONG.length} accepted, no ceiling searched for)`
      : (typeof got === 'string' ? `OBSERVED (came back ${got.length} of ${LONG.length})` : 'FAIL')));

  // ---- L8/L9 -------------------------------------------------------------
  await mergeAndRead('L8', 'Is an EMPTY description accepted?', '',
    (got) => (got === '' ? 'PASS' : `FAIL (read ${JSON.stringify(got)})`));
  await mergeAndRead('L9', 'Does the composed note-plus-marker shape survive intact?', COMPOSED);

  // ---- L10: the one byte run 1 never sent -------------------------------
  // Reports which half survived, so a stripped CR is distinguishable from a failure.
  await mergeAndRead('L10', 'Does a CRLF survive, given a bare LF does?', CRLF_TEXT,
    (got) => {
      if (got === CRLF_TEXT) return 'PASS';
      if (typeof got !== 'string') return 'FAIL';
      if (got.includes(LF) && !got.includes(CR)) return 'FAIL (CR stripped, LF kept)';
      if (got.includes(CR) && !got.includes(LF)) return 'FAIL (LF stripped, CR kept)';
      return 'FAIL';
    });

  await mergeAndRead('L11', 'Does a LONGER run of spaces survive, given two do?',
    LONG_SPACES,
    (got) => (got === LONG_SPACES ? 'PASS'
      : (typeof got === 'string' && got.includes('eight spaces')
        ? 'FAIL (run collapsed)' : 'FAIL')));

  // ---- Clean up ----------------------------------------------------------
  const dd = await getDigest();
  const removed = await spPost(
    `web/lists/getbytitle('${encodeURIComponent(LIST)}')`, {}, dd,
    { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' },
  );
  log(removed.ok ? 'OK' : 'FAIL',
    removed.ok ? `Deleted '${LIST}'.`
      : `Could not delete '${LIST}' (HTTP ${removed.status}). Remove it by hand.`);

  report();
  console.log('');
  console.log('=== WHAT TO DO WITH THIS ===');
  console.log('L4, L5, L6 and L10 together decided entity_note_may_not_round_trip,');
  console.log('which refused an ampersand, a space run, an LF and a CR. All four');
  console.log('must PASS for the rule to stay out. L10 is the only observation of');
  console.log('the CR, so a FAIL there means the rule was right about that byte');
  console.log('even if L4 to L6 pass.');
  console.log('L5 covers ASCII spaces only. Runs containing a tab or a');
  console.log('non-breaking space are still refused by');
  console.log('entity_note_whitespace_unmeasured, because this probe never sent');
  console.log('one. Extend it before allowing those.');
  console.log('L2 is what reconcileListDescription compares on every re-paste, so');
  console.log('a FAIL there is a permanent abort rather than a caveat.');
})();
