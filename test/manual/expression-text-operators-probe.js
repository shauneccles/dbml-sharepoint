/**
 * dbml-sharepoint PROBE: TEXT OPERATORS ON THE EXPRESSION TARGET
 *
 * QUESTION: how should `contains`, `not_contains`, `begins_with` and
 * `not_begins_with` be rendered into a ClientValidationFormula, and does
 * the rendering actually WORK, rather than merely save?
 *
 * WHY: `analysis/conditions.py` disables all four on the expression target:
 *
 *     DISABLED_PENDING_PROBE = { EXPRESSION: _TEXT_OPS }
 *
 * and the build error tells the author to "confirm it with
 * test/manual/form-visibility-evidence-probe.js". That probe does not test
 * them. Its questions are collision, canonical syntax, round-trip fidelity
 * and length limit. So the tool points at a probe that cannot settle the
 * thing it is cited for. This probe is what that error should have named.
 *
 * The validation target already renders all four (`ISNUMBER(FIND(...))`
 * and `LEFT([Col],n)=...`), so this is only about the client-side
 * expression language, which is a different language with different
 * functions.
 *
 * WHY STORAGE IS NOT THE ANSWER, and why this probe ends with an eyes-on
 * checklist you have to fill in: this project has already been caught by
 * `length()`, which is documented, saves cleanly, reads back byte-identical
 * and returns an ARRAY's item count, so `length([$Note]) > 3` is false for
 * every possible value and hides the column unconditionally. A formula that
 * stores perfectly and silently evaluates false is the exact failure this
 * whole surface keeps producing, and no headless script can see it. The
 * console half of this probe narrows the candidates; only your eyes can
 * finish it.
 *
 * WHAT IT ASKS (console half)
 *   X0   NEGATIVE CONTROL: is a formula calling a function that does not
 *        exist REFUSED? If SharePoint stores nonsense, then "accepted"
 *        below means nothing and only the eyes-on half is evidence.
 *   X1   indexOf(...) >= 0            candidate for `contains`
 *   X2   indexOf(...) < 0             candidate for `not_contains`
 *   X3   indexOf(...) == 0            candidate for `begins_with`
 *   X4   startsWith(...)              does this function exist at all?
 *   X5   substring(...) == '...'      documented-functions-only fallback
 *        for `begins_with`, in case X3 or X4 misbehave
 *   X6   indexOf(...) != 0            candidate for `not_begins_with`
 *   Each asks: accepted? and does it read back BYTE-IDENTICAL, or did
 *   SharePoint normalise it (which reconciliation would then have to
 *   canonicalise, as the calculated-formula comparison already does)?
 *
 *   X6 AND THE EMPTY ROW WERE RUN ON 2026-07-29, in a second pass after the
 *   first had settled X1-X5 over three typed values. Both were previously
 *   carried on arithmetic rather than observation: `!= 0` is the exact
 *   negation of the `== 0` watched at X3, and the empty-box behaviour that
 *   `none_of[not_contains]` depends on had never been typed. Result: all
 *   TWENTY-FOUR cells of the six-candidate by four-row table matched the
 *   predicted truth table exactly:
 *
 *       row 1  "needle in a haystack"    -> X1 X3 X4 X5
 *       row 2  "a needle in a haystack"  -> X1 X6
 *       row 3  "no such thing here"      -> X2 X6
 *       row 4  (empty)                   -> X2 X6
 *
 *   X6 is therefore discriminating, not merely stored: hidden for row 1 and
 *   visible for the other three, which is the shape a working candidate has
 *   and neither of the two broken shapes. And row 4 shows BOTH negative
 *   operators visible for the empty box, which is the observation
 *   `analysis/conditions.py` needs to keep none_of[not_contains] from
 *   re-admitting the empty value. The open form, before anything was typed,
 *   independently showed the same two columns, the same answer seen twice.
 *
 *   X1-X5 are unchanged byte-for-byte across both passes, so the first run's
 *   recorded results still describe exactly the strings that produced them.
 *   X0 failed again in the second pass (the nonsense function was stored
 *   byte-identical a second time), so the eyes-on table remains the only
 *   evidence here, and all six candidates storing cleanly means nothing on
 *   its own.
 *
 * WHAT YOU ANSWER (the eyes-on half)
 *   The probe prints a four-row table. You open the New form, put each
 *   value into "ProbeText" (the last row is the box left EMPTY) and write
 *   down which of the six columns appear. A candidate that works shows for
 *   SOME values and not others. A candidate that is broken shows for NONE
 *   (or for ALL), and those two are indistinguishable from a formula that
 *   stored perfectly.
 *
 * HOW TO RUN
 *   1. Open a site you own, at /_layouts/15/settings.aspx.
 *   2. F12 -> Console -> paste -> Enter. It prints its plan and stops.
 *   3. Edit CONFIRMED and ALLOW_WRITES to true, paste again.
 *   4. Copy the RESULTS block back, THEN do the eyes-on checklist it
 *      prints and send that too. Neither half is conclusive alone.
 *
 * WHEN FINISHED: delete the list it created.
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

  const LIST = 'dbmlsp Probe ExprTextOps';
  const SUBJECT = 'ProbeText';
  const NEEDLE = 'needle';
  const fieldsPath = `web/lists/getbytitle('${LIST}')/fields`;

  // Each candidate gets its OWN column, so all five are visible on one form
  // at once and a single pass of the eyes-on table answers everything. One
  // shared column would need five separate runs.
  // X1-X5 are BYTE-IDENTICAL to the first 2026-07-29 run and must stay that
  // way: three places cite this script as the source of that observation, so
  // editing them would leave the citations pointing at something that never
  // ran. X6 was added afterwards and run in a second pass the same day; it
  // came back discriminating (see the header).
  const CANDIDATES = [
    ['X1', 'ShowContains', 'contains',
     `=indexOf([$${SUBJECT}], '${NEEDLE}') >= 0`],
    ['X2', 'ShowNotContains', 'not_contains',
     `=indexOf([$${SUBJECT}], '${NEEDLE}') < 0`],
    ['X3', 'ShowBeginsWith', 'begins_with via indexOf',
     `=indexOf([$${SUBJECT}], '${NEEDLE}') == 0`],
    ['X4', 'ShowStartsWith', 'begins_with via startsWith()',
     `=startsWith([$${SUBJECT}], '${NEEDLE}')`],
    ['X5', 'ShowSubstring', 'begins_with via substring()',
     `=substring([$${SUBJECT}], 0, ${NEEDLE.length}) == '${NEEDLE}'`],
    ['X6', 'ShowNotBeginsWith', 'not_begins_with',
     `=indexOf([$${SUBJECT}], '${NEEDLE}') != 0`],
  ];

  // Chosen so every candidate is discriminated by at least one row:
  //   "needle in a haystack"   contains AND begins with
  //   "a needle in a haystack" contains but does NOT begin with
  //   "no such thing here"     neither
  //   ""                       the empty box
  //
  // The empty row is not decoration. Both NEGATIVE operators are true for a
  // blank (indexOf('', 'needle') is -1, which is `< 0` and also `!= 0`),
  // and `analysis/conditions.py` relies on exactly that to decide that
  // none_of[not_contains] must NOT re-admit the empty value. That was
  // reasoned from the arithmetic, never printed as a row. It is a row now.
  //
  // It is listed last rather than first because the New form opens with the
  // box already empty, so an observer reads it on the way in without
  // typing; putting it at the end keeps the typed rows in one run.
  const CASES = [
    ['needle in a haystack', 'contains AND begins with'],
    ['a needle in a haystack', 'contains, does NOT begin with'],
    ['no such thing here', 'neither'],
    ['', 'THE EMPTY BOX - clear the field rather than typing'],
  ];

  if (!CONFIRMED) {
    log('INFO', `Would create list '${LIST}' on ${WEB} with a text column`);
    log('INFO', `'${SUBJECT}' and ${CANDIDATES.length} more, each carrying one candidate`);
    log('INFO', 'ClientValidationFormula for the disabled text operators,');
    log('INFO', 'then read each back to see whether it survived byte-identical.');
    log('INFO', 'It then prints an EYES-ON checklist you must complete by hand:');
    log('INFO', 'the console half cannot tell a working formula from one that');
    log('INFO', 'stored perfectly and evaluates false for every value.');
    if (CLEANUP) {
      log('INFO', `CLEANUP is ON: '${LIST}' would be RECYCLED first.`);
    } else {
      log('INFO', `CLEANUP is off: columns already on '${LIST}' report "already`);
      log('INFO', 'present" instead of being created. Set CLEANUP = true.');
    }
    log('INFO', 'Nothing has been written. Set CONFIRMED and ALLOW_WRITES to true.');
    return;
  }
  if (!ALLOW_WRITES) {
    log('INFO', 'CONFIRMED, but ALLOW_WRITES is false and this probe must write.');
    log('INFO', 'Set ALLOW_WRITES = true to proceed. Stopping.');
    return;
  }

  expect('X0', 'NEGATIVE CONTROL: a formula calling a non-existent function is refused');
  for (const [id, , label] of CANDIDATES) {
    expect(id, `Candidate for ${label}: accepted, and stored byte-identical?`);
  }

  await resetList(LIST);
  let digest = await getDigest();

  const existing = await spGet(`web/lists/getbytitle('${LIST}')`);
  if (!existing.ok) {
    const made = await spPost('web/lists', {
      Title: LIST,
      BaseTemplate: 100,
      Description: 'dbml-sharepoint probe list. Safe to delete.',
    }, digest);
    if (!made.ok) {
      record('BOOT', 'Create the probe list', 'FAIL',
             `HTTP ${made.status}: ${made.text.slice(0, 300)}`);
      return report();
    }
    log('OK', `Created list '${LIST}'.`);
  }

  const addTextField = async (name) => {
    digest = await getDigest();
    return spPost(`${fieldsPath}/createfieldasxml`, {
      parameters: {
        SchemaXml: `<Field Type="Text" DisplayName="${name}" Name="${name}" />`,
        Options: 8,
      },
    }, digest);
  };
  const fieldExists = async (name) =>
    (await spGet(`${fieldsPath}/getbyinternalnameortitle('${name}')`)).ok;

  for (const name of [SUBJECT, ...CANDIDATES.map((c) => c[1]), 'ShowNegative']) {
    if (!(await fieldExists(name))) {
      const made = await addTextField(name);
      if (!made.ok) {
        record('BOOT', `Create column ${name}`, 'FAIL',
               `HTTP ${made.status}: ${made.text.slice(0, 300)}`);
        return report();
      }
    }
  }
  log('OK', 'Columns ready.');

  const setExpression = async (field, formula) => {
    digest = await getDigest();
    return spPost(`${fieldsPath}/getbyinternalnameortitle('${field}')`,
                  { ClientValidationFormula: formula }, digest,
                  { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
  };
  // Returns the request outcome ALONGSIDE the value. Collapsing a failed
  // GET to null would make a throttled read indistinguishable from a
  // successful read of an empty property, and the classifier below would
  // then print 'ACCEPTED THEN DISCARDED' (a claim about SharePoint) on
  // no observation.
  const readExpression = async (field) => {
    const r = await spGet(
      `${fieldsPath}/getbyinternalnameortitle('${field}')?$select=ClientValidationFormula`);
    return {
      ok: r.ok,
      parsed: r.body !== null,
      status: r.status,
      value: r.ok && r.body ? r.body.ClientValidationFormula : null,
    };
  };

  // ---- X0: NEGATIVE CONTROL -------------------------------------------
  // If SharePoint stores a call to a function that cannot exist, then it is
  // not parsing these formulas at write time at all, and every "ACCEPTED"
  // below is a statement about storage rather than about validity.
  const junkFormula = "=dbmlspNoSuchFunction([$" + SUBJECT + "], 'x')";
  const junk = await setExpression('ShowNegative', junkFormula);
  const junkStored = junk.ok ? (await readExpression('ShowNegative')).value : null;
  record('X0', 'NEGATIVE CONTROL: a formula calling a non-existent function is refused',
         junk.ok ? 'FAIL' : isRefusal(junk.status) ? 'PASS' : 'NOT ESTABLISHED',
         junk.ok
           ? `a call to dbmlspNoSuchFunction was ACCEPTED and stored as `
             + `${JSON.stringify(junkStored)}. SharePoint is not validating these `
             + 'formulas on write, so treat every candidate "ACCEPTED" as storage only '
             + 'and rely entirely on the eyes-on table'
           : isRefusal(junk.status)
             ? `refused with HTTP ${junk.status}: ${junk.text.slice(0, 260)}`
             : `the request failed with HTTP ${junk.status}, which is not the server `
               + 'refusing the formula, so this control established nothing: '
               + junk.text.slice(0, 200));

  // ---- X1-X6: the candidates ------------------------------------------
  for (const [id, field, label, formula] of CANDIDATES) {
    const set = await setExpression(field, formula);
    if (!set.ok) {
      record(id, `Candidate for ${label}: accepted, and stored byte-identical?`,
             'REFUSED', `sent ${JSON.stringify(formula)}; HTTP ${set.status}: `
                        + set.text.slice(0, 260));
      continue;
    }
    const read = await readExpression(field);
    const stored = read.value;
    // Byte-identical matters beyond tidiness: the deploy verifies by
    // read-back, so a formula SharePoint normalises would be reported as
    // drift on every single redeploy unless reconciliation canonicalises
    // it first.
    //
    // Three ways to fail to see a formula, and only one of them is a
    // statement about SharePoint. A failed request and an unparseable body
    // are both NOT ESTABLISHED; a 200 that carried no such property is
    // 'ACCEPTED THEN DISCARDED', which is an observation. `null` and
    // `undefined` are kept apart for the same reason: whether this API
    // spells an absent formula as an explicit null or by omitting the
    // property is not something this repo has established, so the
    // omitted case does not borrow the null case's verdict.
    const outcome = !read.ok || !read.parsed
      ? 'NOT ESTABLISHED'
      : stored === null ? 'ACCEPTED THEN DISCARDED'
      : stored === undefined ? 'NOT ESTABLISHED'
      : stored === formula ? 'ACCEPTED, BYTE-IDENTICAL' : 'ACCEPTED, NORMALISED';
    record(id, `Candidate for ${label}: accepted, and stored byte-identical?`, outcome,
           !read.ok
             ? `sent ${JSON.stringify(formula)}; the MERGE returned HTTP ${set.status} `
               + `but the read-back failed with HTTP ${read.status}`
             : !read.parsed
               ? `sent ${JSON.stringify(formula)}; the read-back returned HTTP `
                 + `${read.status} with a body that did not parse as JSON`
               : stored === undefined
                 ? `sent ${JSON.stringify(formula)}; the read-back carried no `
                   + 'ClientValidationFormula property at all'
                 : `sent ${JSON.stringify(formula)}; stored ${JSON.stringify(stored)}`);
  }

  report();

  // ---- The half a console cannot answer -------------------------------
  console.log('\n============ EYES-ON CHECKLIST (REQUIRED) ============');
  console.log('The rows above say only whether SharePoint KEPT each formula.');
  console.log('They cannot say whether it EVALUATES. Do this now:\n');
  console.log(`  1. Open ${WEB}/Lists/${encodeURIComponent(LIST)}/NewForm.aspx`);
  console.log(`     (or the list -> New). All ${CANDIDATES.length} Show* columns should be`);
  console.log('     on the form; if a column is missing entirely, say so.');
  console.log('     That is a different finding from it being hidden.');
  console.log(`  2. Put each value below into "${SUBJECT}" and write down`);
  console.log('     which Show* columns are VISIBLE. Do not save.\n');
  CASES.forEach(([value, meaning], i) => {
    console.log(`     row ${i + 1}. ${value === '' ? '(leave it EMPTY)' : `"${value}"`}`);
    console.log(`         (${meaning})`);
    console.log('         visible: ______________________________________');
  });
  console.log(`\n  3. Report the ${CASES.length} lines above verbatim.\n`);
  console.log('HOW TO READ IT');
  console.log('  A WORKING candidate is visible for some values and not others:');
  console.log('    contains         -> rows 1 and 2');
  console.log('    not_contains     -> rows 3 and 4');
  console.log('    begins_with      -> row 1 only');
  console.log('    not_begins_with  -> rows 2, 3 and 4');
  console.log('  A BROKEN candidate is visible for ALL rows or for NONE.');
  console.log('  "None" is the dangerous one: it looks exactly like a formula');
  console.log('  that stored perfectly, which is how length() fooled this');
  console.log('  project once already.\n');
  console.log('  ROW 4 IS THE ONE TO CHECK CAREFULLY. Both negative operators');
  console.log('  must be VISIBLE for the empty box, because indexOf on an empty');
  console.log('  string is -1. conditions.py depends on that: it is why');
  console.log('  none_of[not_contains] does not re-admit the empty value. If');
  console.log('  row 4 hides X2 or X6, that assumption is wrong and the');
  console.log('  none_of normaliser has to change.');
  console.log('======================================================');
  log('INFO', `Done. Delete '${LIST}' when you have finished the checklist.`);
})();
