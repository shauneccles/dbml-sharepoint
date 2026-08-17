/**
 * dbml-sharepoint PROBE: WHAT A LOOKUP SHOWS, AND WHAT IT OFFERS
 *
 * TWO QUESTIONS about the same column type, asked together because they
 * need the same pair of lists.
 *
 * QUESTION A: DOES A LOOKUP LEAK ACROSS AN ACL? When list A carries a
 * Lookup into list B, and a reader is DENIED list B, does the referencing
 * item on list A still show them B's display value?
 *
 * WHY: service-evidence-register puts `RelatedIssue` on `ServiceEvent`,
 * which every contributor can read, pointing at `ServiceIssue`, which they
 * are deliberately denied. 50-govern says "Contributors deliberately
 * cannot see ServiceIssue... a leaked one is worse than no register at
 * all". A review raised that the lookup may render the theme's title on
 * the event anyway, defeating the boundary. Nobody has measured it. This
 * project has been wrong about SharePoint by reasoning from plausibility
 * before, so the register ships the guidance that is safe under both
 * answers and this file exists to settle which answer it is.
 *
 * QUESTION B: CAN A LOOKUP OFFER FEWER ROWS? A lookup picker lists every
 * item in the target, which stops being usable at a few hundred rows and
 * offers choices that are wrong rather than merely many: a closed theme
 * should not be selectable for a new event. The widely-repeated remedy is
 * to point the lookup at a CALCULATED column that returns an empty string
 * for the rows you want hidden, on the belief that the picker omits rows
 * whose display value is empty. Two things about that are unverified here
 * and one of them can lose data:
 *
 *   - does SharePoint accept a calculated column as a lookup's display
 *     field at all? The tool's own validator does not check the TYPE of
 *     `display_column`, only that the name exists, so it would emit
 *     LookupField pointing at one and find out at deploy time.
 *   - what happens to rows ALREADY LINKED when their label later becomes
 *     empty? If an event links to a theme and the theme is then closed,
 *     the trick would blank the link on every event behind it. That is a
 *     worse failure than a long picker, and it is the reason this asks
 *     rather than recommends.
 *
 * WHAT IT ASKS
 *   -- Question A, needs the second account ------------------------------
 *   K1   CONTROL: is the second account actually denied the TARGET list?
 *        If it can still open the target, every row below means nothing.
 *   K2   THE ONE THAT MATTERS. Reading the SOURCE item as the denied
 *        account, does the lookup's display value come back?
 *   K3   ...and does $expand reach the target row's OTHER columns, or only
 *        the display field? A denied reader who can expand is a bigger
 *        hole than one who sees a title.
 *   K4   CONTROL: can the denied account read the SOURCE list at all? If
 *        not, K2 and K3 are silent for the wrong reason.
 *
 *   -- Question B, answered by the site owner ----------------------------
 *   K5   does SharePoint accept a CALCULATED column as a lookup's display
 *        field (LookupFieldName)?
 *   K6   ...and with the label empty for a row, what does an item ALREADY
 *        LINKED to that row read back as? This is the risky one: a theme that
 *        closes must not blank the link on the events behind it.
 *   K7   EYES-ON: does the New form's picker actually omit the row whose
 *        calculated label is empty? A picker is a rendering surface and no
 *        REST call can answer it.
 *
 * READ K1 AND K4 FIRST. K2 is evidence only when the account is provably
 * denied the target AND provably allowed the source.
 *
 * TWO ACCOUNTS, TWO PASTES (question A only). No probe here has needed a
 * second identity before, and a site collection administrator cannot deny
 * themselves, so running it all as one person would answer a question
 * nobody asked.
 *
 *   PASS 1: as a SITE OWNER. MODE = 'setup'. Creates two lists, links
 *            rows, answers K5 and K6, prints the K7 checklist, then breaks
 *            inheritance on the target and strips every assignment except
 *            Site Owners.
 *   PASS 2: as a SECOND, NON-PRIVILEGED account, a member or visitor of
 *            this site who is NOT a site owner, site collection admin or
 *            tenant admin. MODE = 'read'. Writes NOTHING.
 *
 * If the second account turns out to be an administrator, K1 says so
 * rather than letting the run look successful.
 *
 * HOW TO RUN
 *   1. As the site owner, open a site you own at /_layouts/15/settings.aspx.
 *   2. F12 -> Console -> paste -> Enter. It prints its plan and stops.
 *   3. Set CONFIRMED and ALLOW_WRITES to true, leave MODE = 'setup', paste
 *      again. Do the K7 checklist it prints before signing out.
 *   4. Sign in as the second account (a private window is easiest), open
 *      the same settings page, set MODE = 'read', paste. CONFIRMED alone is
 *      enough for pass 2; it never writes.
 *   5. Copy BOTH results blocks back, and the eyes-on lines.
 *
 * WHEN FINISHED: delete both lists as the site owner.
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

  // 'setup' writes and needs a site owner; 'read' only reads and is the
  // pass that answers question A. Committed as 'setup' so a stray paste by
  // an owner is the harmless half.
  const MODE = 'setup';

  const TARGET = 'dbmlsp Probe LookupTarget';
  const SOURCE = 'dbmlsp Probe LookupSource';
  const LOOKUP = 'ProbeLink';
  const PICK = 'ProbePick';
  const LABEL = 'ProbePickLabel';
  // Distinctive enough that finding it in a response is unambiguous, and
  // obviously not real data to anyone who stumbles on the list.
  const SECRET = 'dbmlsp-probe-target-title-should-not-leak';
  const SIDE = 'dbmlsp-probe-target-second-column';
  const CLOSED_TITLE = 'dbmlsp-probe-closed-row';

  expect('K1', 'CONTROL: is the second account actually denied the TARGET list?');
  expect('K2', 'Reading the SOURCE item as the denied account, does the lookup display value come back?');
  expect('K3', 'Does $expand on the lookup reach the target row\'s other columns?');
  expect('K4', 'CONTROL: can the denied account read the SOURCE list at all?');
  expect('K5', 'Does SharePoint accept a CALCULATED column as a lookup display field?');
  expect('K6', 'With the label empty, what does an ALREADY LINKED item read back as?');
  expect('K7', 'EYES-ON: does the picker omit the row whose calculated label is empty?');

  if (!CONFIRMED) {
    log('INFO', `MODE is '${MODE}'.`);
    if (MODE === 'setup') {
      log('INFO', `Would create lists '${TARGET}' and '${SOURCE}' on ${WEB}, add a`);
      log('INFO', `lookup '${LOOKUP}' on Title and a second lookup '${PICK}' on a`);
      log('INFO', `calculated column '${LABEL}' that is empty for a closed row,`);
      log('INFO', 'create three rows and link them, then BREAK INHERITANCE on the');
      log('INFO', 'target and remove every assignment except Site Owners.');
      log('INFO', 'Nothing else on the site is touched.');
      if (CLEANUP) {
        log('INFO', `CLEANUP is ON: '${TARGET}' and '${SOURCE}' would be RECYCLED first.`);
      }
    } else {
      log('INFO', 'Would only READ, as the account you are signed in as now.');
      log('INFO', 'Run this pass as a NON-PRIVILEGED account, not the site owner.');
    }
    log('INFO', 'Nothing has been written. Set CONFIRMED (and ALLOW_WRITES for setup).');
    return;
  }

  // ---- PASS 2: read as the denied account -----------------------------
  if (MODE === 'read') {
    const me = await spGet('web/currentuser?$select=Title,LoginName,IsSiteAdmin');
    const who = me.ok && me.body
      ? `${me.body.Title} (${me.body.LoginName}), IsSiteAdmin=${me.body.IsSiteAdmin}`
      : `could not read web/currentuser (HTTP ${me.status})`;
    log('INFO', `Running as: ${who}`);

    // K1 is the control. A site collection admin is never denied anything,
    // so an admin running this pass invalidates the run and must be told,
    // not quietly passed.
    const target = await spGet(`web/lists/getbytitle('${TARGET}')/items?$select=Title&$top=5`);
    const isAdmin = me.ok && me.body && me.body.IsSiteAdmin === true;
    // Only 401/403 is a DENIAL. A 429 or a 500 also fails, and reading
    // either as "denied" would let K2 run believing a premise it has not
    // established. The throttled case would then report the lookup value
    // withheld from an account that was never actually refused anything.
    const deniedTarget = target.status === 401 || target.status === 403;
    if (!me.ok) {
      record('K1', 'CONTROL: is the second account actually denied the TARGET list?',
             'NOT ESTABLISHED',
             `could not read web/currentuser (HTTP ${me.status}), so this run cannot `
             + 'even say who it is running as, let alone whether they are privileged.');
    } else if (isAdmin) {
      record('K1', 'CONTROL: is the second account actually denied the TARGET list?',
             'NOT ESTABLISHED',
             `this account is a site collection administrator (${who}). SharePoint `
             + 'does not apply broken inheritance to one, so nothing below can be '
             + 'read as evidence. Re-run pass 2 as a non-privileged account.');
    } else if (target.ok) {
      record('K1', 'CONTROL: is the second account actually denied the TARGET list?',
             'FAIL',
             `the account READ ${TARGET} (HTTP ${target.status}), so inheritance was `
             + 'not broken as intended and K2/K3 prove nothing about a denied reader');
    } else if (!deniedTarget) {
      record('K1', 'CONTROL: is the second account actually denied the TARGET list?',
             'NOT ESTABLISHED',
             `the target read failed with HTTP ${target.status}, which is not an access `
             + 'denial. A throttled or erroring request is not evidence that the ACL '
             + 'holds. Re-run.');
    } else {
      record('K1', 'CONTROL: is the second account actually denied the TARGET list?',
             'PASS',
             `refused with HTTP ${target.status}, so the account is denied the `
             + `target list. Running as: ${who}`);
    }

    // K4 is the other control, asked before K2 so the read-out order is the
    // order a reader needs them in.
    const source = await spGet(
      `web/lists/getbytitle('${SOURCE}')/items?$select=Title,${LOOKUP}Id&$top=10`);
    const rows = (source.ok && source.body && source.body.value) || [];
    // Reading the list is not enough. K2 concludes "withheld" from the
    // ABSENCE of a string, so the row that would carry it has to be proven
    // present and linked first. Otherwise a deleted fixture, or one past
    // the page limit, reads as a clean security result.
    const fixture = rows.find((r) => r.Title === 'dbmlsp-probe-source-row');
    const fixtureLinked = Boolean(fixture && fixture[`${LOOKUP}Id`]);
    record('K4', 'CONTROL: can the denied account read the SOURCE list at all?',
           !source.ok ? 'FAIL' : fixtureLinked ? 'PASS' : 'NOT ESTABLISHED',
           !source.ok
             ? `refused with HTTP ${source.status}. K2 and K3 are silent for the `
               + 'wrong reason. Grant this account read on the source list and re-run.'
             : fixtureLinked
               ? `read ${rows.length} row(s), including the linked fixture row `
                 + `(${LOOKUP}Id=${fixture[`${LOOKUP}Id`]})`
               : `read ${rows.length} row(s) from ${SOURCE}, but the linked fixture row `
                 + `'dbmlsp-probe-source-row' is ${fixture ? 'present without a lookup id'
                                                          : 'not among them'}. K2 would `
                 + 'be reading the absence of a string that was never going to be there. '
                 + 'Re-run the setup pass.');

    if (!me.ok || !source.ok || !fixtureLinked || isAdmin || target.ok || !deniedTarget) {
      record('K2', 'Reading the SOURCE item as the denied account, does the lookup display value come back?',
             'NOT ESTABLISHED', 'a control above did not hold (see K1 and K4)');
      record('K3', 'Does $expand on the lookup reach the target row\'s other columns?',
             'NOT ESTABLISHED', 'a control above did not hold (see K1 and K4)');
    } else {
      const expanded = await spGet(
        `web/lists/getbytitle('${SOURCE}')/items?$select=Title,${LOOKUP}/Title`
        + `&$expand=${LOOKUP}&$top=10`);
      const raw = expanded.ok ? JSON.stringify(expanded.body) : '';
      const denied = expanded.status === 401 || expanded.status === 403;
      if (!expanded.ok && !denied) {
        record('K2', 'Reading the SOURCE item as the denied account, does the lookup display value come back?',
               'NOT ESTABLISHED',
               `the $expand read failed with HTTP ${expanded.status}, which is neither `
               + 'a denial nor an answer');
      } else {
        const leaked = raw.includes(SECRET);
        record('K2', 'Reading the SOURCE item as the denied account, does the lookup display value come back?',
               leaked ? 'LOOKUP VALUE IS VISIBLE' : 'LOOKUP VALUE IS WITHHELD',
               leaked
                 ? 'the target Title came back to an account denied the target list: '
                   + `the response contains ${JSON.stringify(SECRET)}. A lookup does not `
                   + 'respect the target list ACL, and any register relying on one to '
                   + 'hide a title is relying on nothing.'
                 : `HTTP ${expanded.status}; the response does not contain `
                   + `${JSON.stringify(SECRET)}. Body: ${JSON.stringify(raw.slice(0, 300))}`);
      }

      const side = await spGet(
        `web/lists/getbytitle('${SOURCE}')/items?$select=${LOOKUP}/Title,${LOOKUP}/ProbeSide`
        + `&$expand=${LOOKUP}&$top=10`);
      const sideRaw = side.ok ? JSON.stringify(side.body) : '';
      // A refusal is the server rejecting the query or the access. A 429 or
      // a 500 is neither, and calling one REFUSED would turn a throttle into
      // an ACL result, which K1 and K2 above already avoid.
      const sideRefused = isRefusal(side.status)
        || side.status === 401 || side.status === 403;
      record('K3', 'Does $expand on the lookup reach the target row\'s other columns?',
             side.ok
               ? sideRaw.includes(SIDE) ? 'OTHER COLUMNS ALSO VISIBLE'
                 : sideRaw.includes(SECRET) ? 'DISPLAY FIELD ONLY'
                 : 'NOT ESTABLISHED'
               : sideRefused ? 'REFUSED' : 'NOT ESTABLISHED',
             side.ok
               ? sideRaw.includes(SIDE)
                 ? 'a column that is NOT the lookup display field came back too '
                   + `(${JSON.stringify(SIDE)}), so the exposure is wider than the title`
                 : sideRaw.includes(SECRET)
                   ? `the display field came back and ${JSON.stringify(SIDE)} did not`
                   : 'NEITHER field came back, so this says nothing about the second '
                     + 'one: the response may simply not carry the linked row. Body: '
                     + JSON.stringify(sideRaw.slice(0, 300))
               : sideRefused
                 ? '$expand naming a second target column was refused with HTTP '
                   + `${side.status}: ${JSON.stringify(String(side.body).slice(0, 200))}`
                 : `the $expand request failed with HTTP ${side.status}, which is `
                   + 'neither a refusal nor an answer');
    }

    report();
    console.log('\n============ EYES-ON, PASS 2 ============');
    console.log('REST is not the only surface. Still signed in as this account:');
    console.log(`  1. Open ${WEB}/Lists/${encodeURIComponent(SOURCE)}/AllItems.aspx`);
    console.log(`  2. Look at the '${LOOKUP}' column, and open an item.`);
    console.log('     Does the target row title appear, is the cell blank, or is');
    console.log('     there an error?');
    console.log('     what you see: ______________________________________');
    console.log('==========================================');
    log('INFO', 'Read-only pass complete. Nothing was written.');
    return;
  }

  // ---- PASS 1: setup, as the site owner -------------------------------
  if (!ALLOW_WRITES) {
    log('INFO', 'CONFIRMED, but ALLOW_WRITES is false and setup must write.');
    log('INFO', 'Set ALLOW_WRITES = true to proceed. Stopping.');
    return;
  }

  await resetList(SOURCE);
  await resetList(TARGET);
  let digest = await getDigest();

  const bail = (id, question, why) => {
    record(id, question, 'NOT ESTABLISHED', why);
    report();
  };

  const ensureList = async (title) => {
    const found = await spGet(`web/lists/getbytitle('${title}')`);
    if (found.ok) {
      log('INFO', `List '${title}' already exists, reusing it.`);
      return found.body;
    }
    digest = await getDigest();
    const made = await spPost('web/lists', {
      Title: title,
      BaseTemplate: 100,
      Description: 'dbml-sharepoint probe list. Safe to delete.',
    }, digest);
    if (!made.ok) {
      log('FAIL', `Could not create '${title}': HTTP ${made.status} ${made.text.slice(0, 240)}`);
      return null;
    }
    log('OK', `Created list '${title}'.`);
    return made.body;
  };

  const addField = async (list, schemaXml) => {
    digest = await getDigest();
    return spPost(`web/lists/getbytitle('${list}')/fields/createfieldasxml`, {
      parameters: { SchemaXml: schemaXml, Options: 8 },
    }, digest);
  };

  const targetList = await ensureList(TARGET);
  if (!targetList) {
    return bail('K1', 'CONTROL: is the second account actually denied the TARGET list?',
                `the target list could not be created (see the FAIL above)`);
  }
  const sourceList = await ensureList(SOURCE);
  if (!sourceList) {
    return bail('K1', 'CONTROL: is the second account actually denied the TARGET list?',
                `the source list could not be created (see the FAIL above)`);
  }

  // A second target column, so K3 can ask whether $expand reaches past the
  // display field; and a status column the calculated label keys on.
  await addField(TARGET, '<Field Type="Text" DisplayName="ProbeSide" Name="ProbeSide" />');
  await addField(TARGET, '<Field Type="Text" DisplayName="ProbeStatus" Name="ProbeStatus" />');

  // The label the picker would be pointed at: empty for a closed row.
  const labelXml =
    `<Field Type="Calculated" DisplayName="${LABEL}" Name="${LABEL}" ResultType="Text">`
    + '<Formula>=IF([ProbeStatus]="Closed","",[Title])</Formula>'
    + '<FieldRefs><FieldRef Name="ProbeStatus"/><FieldRef Name="Title"/></FieldRefs>'
    + '</Field>';
  const labelMade = await addField(TARGET, labelXml);
  if (!labelMade.ok) {
    log('FAIL', `Calculated label column refused: HTTP ${labelMade.status} `
                + labelMade.text.slice(0, 240));
  } else {
    log('OK', `Added calculated column '${LABEL}'.`);
  }

  digest = await getDigest();
  const openRow = await spPost(`web/lists/getbytitle('${TARGET}')/items`,
                               { Title: SECRET, ProbeSide: SIDE, ProbeStatus: 'Open' }, digest);
  digest = await getDigest();
  // Created OPEN, so its calculated label is populated. K6 closes it AFTER
  // the link exists. The transition is the question, and a row born closed
  // would only have shown what linking to an already-empty label does.
  const closedRow = await spPost(`web/lists/getbytitle('${TARGET}')/items`,
                                 { Title: CLOSED_TITLE, ProbeStatus: 'Open' }, digest);
  if (!openRow.ok || !closedRow.ok) {
    return bail('K1', 'CONTROL: is the second account actually denied the TARGET list?',
                `could not create the target rows: HTTP ${openRow.status}/${closedRow.status}`);
  }
  const openId = openRow.body.Id;
  const closedId = closedRow.body.Id;
  log('OK', 'Created one open and one closed row on the target.');

  const targetGuid = targetList.Id;
  const addLookup = async (name, lookupFieldName) => {
    digest = await getDigest();
    return spPost(`web/lists/getbytitle('${SOURCE}')/fields/addfield`, {
      parameters: {
        Title: name,
        FieldTypeKind: 7,
        LookupListId: targetGuid,
        LookupFieldName: lookupFieldName,
      },
    }, digest);
  };

  // The ordinary lookup, on Title. This is the one question A is about.
  const plainLookup = await addLookup(LOOKUP, 'Title');
  if (!plainLookup.ok) {
    return bail('K1', 'CONTROL: is the second account actually denied the TARGET list?',
                `could not add the lookup column: HTTP ${plainLookup.status} `
                + plainLookup.text.slice(0, 240));
  }
  log('OK', `Added lookup '${LOOKUP}' -> ${TARGET}.Title.`);

  // K5 is the calculated display field. The tool's validator would let this
  // through (it checks the NAME exists, never the type), so whether the
  // platform accepts it is the whole question.
  const calcLookup = await addLookup(PICK, LABEL);
  // The calculated label column has to exist before any of this means
  // anything: if addField refused IT, addfield refusing the lookup says
  // nothing about calculated display fields.
  record('K5', 'Does SharePoint accept a CALCULATED column as a lookup display field?',
         !labelMade.ok ? 'NOT ESTABLISHED'
           : calcLookup.ok ? 'ACCEPTED'
           : isRefusal(calcLookup.status) ? 'REFUSED' : 'NOT ESTABLISHED',
         !labelMade.ok
           ? `the calculated column '${LABEL}' could not be created (HTTP `
             + `${labelMade.status}), so there was never a calculated display field to `
             + 'point a lookup at: ' + labelMade.text.slice(0, 200)
           : calcLookup.ok
             ? `addfield with LookupFieldName='${LABEL}' (a Calculated column) returned `
               + `HTTP ${calcLookup.status}. Accepted is not the same as usable. K6 and `
               + 'K7 are what decide that.'
             : isRefusal(calcLookup.status)
               ? `HTTP ${calcLookup.status}: ${calcLookup.text.slice(0, 300)}. The `
                 + 'empty-label trick cannot be built this way on this tenant, so a '
                 + 'shorter picker needs a different mechanism.'
               : `the request failed with HTTP ${calcLookup.status}, which is not the `
                 + 'server refusing the column: ' + calcLookup.text.slice(0, 200));

  digest = await getDigest();
  const sourceRow = await spPost(`web/lists/getbytitle('${SOURCE}')/items`, {
    Title: 'dbmlsp-probe-source-row',
    [`${LOOKUP}Id`]: openId,
  }, digest);
  if (!sourceRow.ok) {
    return bail('K1', 'CONTROL: is the second account actually denied the TARGET list?',
                `could not create the linked source row: HTTP ${sourceRow.status} `
                + sourceRow.text.slice(0, 240));
  }
  log('OK', 'Created and linked one source row.');

  // K6 is the risk, and the ORDER is the whole point. Link a row through the
  // CALCULATED lookup while the target's label is still populated, THEN
  // close the target so the label empties, THEN read the link back. If it
  // comes back blank, closing a theme would blank the link on every event
  // behind it, and the trick costs history to buy a shorter picker.
  //
  // Linking to a row that was already empty would answer a different and
  // much less interesting question.
  if (!calcLookup.ok) {
    record('K6', 'With the label empty, what does an ALREADY LINKED item read back as?',
           'NOT ESTABLISHED', 'the calculated lookup column was refused at K5');
    record('K7', 'EYES-ON: does the picker omit the row whose calculated label is empty?',
           'NOT ESTABLISHED', 'the calculated lookup column was refused at K5');
  } else {
    digest = await getDigest();
    const linkedToClosed = await spPost(`web/lists/getbytitle('${SOURCE}')/items`, {
      Title: 'dbmlsp-probe-linked-to-closed',
      [`${PICK}Id`]: closedId,
    }, digest);
    if (!linkedToClosed.ok) {
      record('K6', 'With the label empty, what does an ALREADY LINKED item read back as?',
             'NOT ESTABLISHED',
             `could not link a row through '${PICK}' to the target row: HTTP `
             + `${linkedToClosed.status}: ${linkedToClosed.text.slice(0, 240)}`);
    } else {
      // BEFORE closing: prove the label is actually rendering. K6's whole
      // conclusion is drawn from the title's ABSENCE afterwards, so without
      // this a calculated column that never computed at all would read as
      // "the transition blanked it". The absence has to be shown to be a
      // CHANGE rather than the way it always was.
      const before = await spGet(
        `web/lists/getbytitle('${SOURCE}')/items(${linkedToClosed.body.Id})`
        + `?$select=${PICK}/${LABEL}&$expand=${PICK}`);
      const labelWasRendering =
        !readFailed(before) && JSON.stringify(before.body).includes(CLOSED_TITLE);
      if (!labelWasRendering) {
        record('K6', 'With the label empty, what does an ALREADY LINKED item read back as?',
               'NOT ESTABLISHED',
               `the link was made, but the label was NOT rendering before the target `
               + `was closed (HTTP ${before.status}; body `
               + `${JSON.stringify(String(JSON.stringify(before.body)).slice(0, 200))}). `
               + 'There is no transition to observe: an empty read afterwards would be '
               + 'the state it started in.');
        return report();
      }
      log('OK', `'${LABEL}' renders through the lookup before closing.`);

      // Now close it. The label goes empty UNDER an existing link, which is
      // the state a theme reaches when a curator concludes it.
      digest = await getDigest();
      const closed = await spPost(
        `web/lists/getbytitle('${TARGET}')/items(${closedId})`,
        { ProbeStatus: 'Closed' }, digest,
        { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
      if (!closed.ok) {
        record('K6', 'With the label empty, what does an ALREADY LINKED item read back as?',
               'NOT ESTABLISHED',
               `the link was made, but closing the target failed (HTTP ${closed.status}), `
               + `so the label never emptied and there is no transition to observe: `
               + closed.text.slice(0, 200));
        return report();
      }
      log('OK', `Closed the target row, so '${LABEL}' is now empty under a live link.`);
      const readBack = await spGet(
        `web/lists/getbytitle('${SOURCE}')/items(${linkedToClosed.body.Id})`
        + `?$select=Title,${PICK}Id,${PICK}/${LABEL}&$expand=${PICK}`);
      if (!readBack.ok || !readBack.body) {
        record('K6', 'With the label empty, what does an ALREADY LINKED item read back as?',
               'NOT ESTABLISHED',
               `the read-back failed with HTTP ${readBack.status}, so this run has no `
               + 'evidence either way');
      } else {
        const body = JSON.stringify(readBack.body);
        const keepsId = body.includes(`"${PICK}Id"`);
        const showsTitle = body.includes(CLOSED_TITLE);
        record('K6', 'With the label empty, what does an ALREADY LINKED item read back as?',
               showsTitle ? 'LABEL STILL RENDERS' : 'LABEL READS EMPTY',
               `${showsTitle
                   ? 'the closed row title came back anyway'
                   : 'the display value is empty for a row that IS linked'}`
               + `${keepsId ? ' (the id is still stored)' : ' (no id came back either)'}`
               + `. Body: ${JSON.stringify(body.slice(0, 300))}`);
      }
    }
  }

  // Break inheritance on the TARGET and strip every assignment except the
  // owners group, the same shape `list_permissions` deploys, done by hand
  // because this probe asks what that shape actually buys.
  digest = await getDigest();
  const broke = await spPost(
    `web/lists/getbytitle('${TARGET}')/breakroleinheritance(copyRoleAssignments=false,clearSubscopes=true)`,
    {}, digest);
  if (!broke.ok) {
    return bail('K1', 'CONTROL: is the second account actually denied the TARGET list?',
                `breakroleinheritance failed: HTTP ${broke.status} ${broke.text.slice(0, 240)}`);
  }
  log('OK', `Broke inheritance on '${TARGET}' with NO copied assignments.`);

  const owners = await spGet('web/associatedownergroup?$select=Id,Title');
  if (owners.ok && owners.body && owners.body.Id) {
    digest = await getDigest();
    const granted = await spPost(
      `web/lists/getbytitle('${TARGET}')/roleassignments/addroleassignment`
      + `(principalid=${owners.body.Id},roledefid=1073741829)`, {}, digest);
    log(granted.ok ? 'OK' : 'FAIL',
        granted.ok
          ? `Granted Full Control on '${TARGET}' to '${owners.body.Title}' only.`
          : `Could not grant owners on '${TARGET}': HTTP ${granted.status}`);
  } else {
    log('FAIL', 'Could not read the associated owner group. Grant yourself access '
                + `to '${TARGET}' by hand before deleting it.`);
  }

  report();
  console.log('\n============ EYES-ON, PASS 1 (K7) ============');
  console.log('Do this BEFORE signing out; a picker is a rendering surface and');
  console.log('no REST call can answer it.');
  console.log(`  1. Open ${WEB}/Lists/${encodeURIComponent(SOURCE)}/NewForm.aspx`);
  console.log(`  2. Open the '${PICK}' picker, the one on the CALCULATED label.`);
  console.log(`     Rows on the target are "${SECRET}" (label populated) and`);
  console.log(`     "${CLOSED_TITLE}" (label empty, because it is Closed).`);
  console.log('     Which rows does the picker offer?');
  console.log('     offered: ______________________________________');
  console.log(`  3. Open the '${LOOKUP}' picker (the one on Title) for contrast.`);
  console.log('     offered: ______________________________________');
  console.log('');
  console.log('  A picker that omits the empty-label row is the mechanism for a');
  console.log('  shorter, correct choice list. One that offers a BLANK entry is');
  console.log('  worse than doing nothing: the choice is still there and now has');
  console.log('  no name. Read this together with K6: if a linked row goes blank');
  console.log('  when its label empties, the trick costs history to buy tidiness.');
  console.log('');
  console.log('============ NOW RUN PASS 2 ============');
  console.log('Question A is unanswered until a SECOND account runs the read');
  console.log('pass. That account must NOT be a site owner, site collection');
  console.log('administrator or tenant administrator. K1 checks.');
  console.log('');
  console.log(`  1. Give the second account read access to '${SOURCE}' only.`);
  console.log(`     It must have NOTHING on '${TARGET}'.`);
  console.log('  2. Sign in as that account (a private window is easiest) and');
  console.log(`     open ${WEB}/_layouts/15/settings.aspx`);
  console.log("  3. Paste this same file with MODE = 'read' and CONFIRMED = true.");
  console.log('  4. Copy back BOTH results blocks and the eyes-on lines.');
  console.log('==========================================');
  log('INFO', `Delete '${SOURCE}' and '${TARGET}' when you have finished.`);
})();
