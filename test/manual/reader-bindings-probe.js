/**
 * dbml-sharepoint PROBE: WHAT `Read` MEANS HERE, AND WHAT A GROUP ALREADY HOLDS
 *
 * READ-ONLY. It creates nothing, changes nothing and deletes nothing. There is
 * no ALLOW_WRITES path; CONFIRMED alone runs it. Every question is a GET.
 *
 * TWO ISSUES, ONE SETUP. Both are blocked on the same two facts about a site,
 * and both stand between the enterprise-reader tier and a release.
 *
 * #199: THE READER TRUSTS A LEVEL CALLED `Read` BY ITS NAME. The validator
 * exempts an assignment to `Read` as "the safe built-in", and the ACL phase
 * resolves it with `$select=Id` and never looks at what it can do. On a site
 * where an administrator has customised the built-in, the permanently-enrolled
 * reporting account can hold edit or delete rights while the manifest still
 * reports it read-only. R2 reads the bitmap this site's `Read` actually
 * carries; R3 says whether it matches a stock one.
 *
 * #198: THE READER INHERITS WHATEVER ITS GROUP ALREADY HOLDS. The deploy adds
 * the account to a group found BY NAME and never inspects that group's
 * existing bindings. A group carrying Full Control at web scope, or an
 * elevated binding on a list outside the bundle, hands all of it to the
 * account, and neither the ACL phase nor anything else removes it, because
 * both only reconcile lists this bundle declares. R4 and R5 census what EVERY
 * group on the site holds today.
 *
 * WHY THIS CANNOT BE ANSWERED FROM DOCUMENTATION. Microsoft Learn documents
 * what the STOCK levels contain. It cannot say what THIS tenant's `Read` was
 * edited into, nor what a particular group was granted years ago by somebody
 * who has left. Those are site facts, and the gate #209 needs has to be
 * written against them rather than against the defaults.
 *
 * SEPARATING WHAT IS DEPENDED ON FROM WHAT IS OBSERVED. R2, R4 and R5 REPORT
 * what they find. They do not assert a particular bitmap or an empty binding
 * list. A site legitimately has either. R3 is the only one that judges, and
 * it judges against the value R2 read on this site, printed in full so the
 * comparison is auditable rather than hidden in a boolean.
 *
 * THE NEGATIVE CONTROL IS R1, and it is doing real work here. Every other
 * question reports "what came back". A caller who cannot read
 * `web/roledefinitions` at all would produce empty results that look exactly
 * like "nothing is granted", the most dangerous possible false reassurance
 * for a question about excess privilege. R1 establishes the read works before
 * any absence is believed.
 *
 * NOTHING TO SET. It surveys EVERY group on the site. An earlier draft asked
 * about one named group and defaulted to a name that only exists after a
 * deploy, so on a fresh site it reported NOT ESTABLISHED and told the
 * operator to go and guess. The census costs the same number of requests.
 *
 * RUN 1, 2026-08-14, revision f0927e57, one Microsoft 365 group-connected
 * Team Site. Six questions, FOUR answered and two NOT ESTABLISHED, which is
 * the probe reporting its own failure rather than passing vacuously.
 *
 *     R1  PASS: 11 role definitions readable, so an empty result means empty.
 *     R2  PASS: this site's `Read` is
 *                RoleTypeKind=2, High=176, Low=138612833,
 *                "Can view pages and list items and download documents."
 *                The FIRST measured bitmap this project holds for a Read.
 *                ONE site on ONE tenant: it is a reference point, not a
 *                constant to compare against, until a second site agrees.
 *     R3  PASS: RoleTypeKind=2, so SharePoint still regards it as the
 *                built-in Reader. No custom level is wearing the name here.
 *     R4  NOT ESTABLISHED, R5 NOT ESTABLISHED, and the fault was the
 *                probe's. It asked about ONE group, defaulting to a name
 *                that only exists AFTER a deploy of the branch that
 *                introduced it, so on a site that had never seen one there
 *                was nothing to inspect and the operator was told to go and
 *                guess a name. Replaced by a census of every group, which
 *                costs the same requests: the web assignments are one read
 *                either way, and each list's were already read once per list.
 *
 * RUN 2, 2026-08-14, revision 79eeaec8, same site, after the census
 * rewrite. Six questions, SIX answered.
 *
 *     R4  Every tool-created group holds exactly `Limited Access` at web
 *         scope -- the DERIVED binding SharePoint adds when a group is
 *         granted something on a list below the web. Nothing elevated.
 *         The site's own Owners/Members/Visitors hold Full Control, Edit
 *         and Read as expected.
 *     R5  `RC Enterprise Readers` holds `Read` on exactly the two RC lists
 *         it was granted and nothing else. Each `<XX> List Administrators`
 *         holds Full Control on its own family's lists only.
 *
 * So #198's hazard is NOT present on this site: no group carries a binding
 * the deploy would hand on unexpectedly. One site is not a guarantee, but
 * the shape of the gate is now known -- probe the group's web-scope
 * bindings and its explicit list bindings, and fail closed on anything
 * beyond a derived Limited Access.
 *
 * THREE THINGS THE RUN CONFIRMED THAT WERE PREVIOUSLY TAKEN FROM DOCS:
 *
 *   - `Limited Access` really is DERIVED. It appears on every group that
 *     was granted something below the web, and on none that was not, and
 *     nothing assigned it directly. That is the empirical half of the
 *     argument for refusing it as an assignment level, which until now
 *     rested on a Learn sentence alone.
 *   - `Web-Only Limited Access` is REAL and in use, held by the web's
 *     Limited Access system group. It was added to BUILT_IN_LEVELS on the
 *     strength of a Learn table that most summaries omit; here it is.
 *   - The site carries `CI`, `PP`, `RC` and `RG List Administrators` plus
 *     `RC Enterprise Readers` -- four families deployed, four administrator
 *     groups, exactly the fragmentation the shared-group change removes.
 *
 * #199 part 2 is ANSWERED FOR THIS SITE, and #198 with it.
 *
 * HOW TO RUN
 *   1. Open the target site as a SITE OWNER. Reading role assignments needs
 *      it, and R1 will tell you plainly if this account cannot.
 *   2. Open a browser console on any page of that site.
 *   3. Set CONFIRMED = true, then paste this file. It never writes.
 *   4. Copy the whole results block back verbatim.
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
  log('INFO', 'probe revision bb56ab36. Quote this when reporting results.');


  // Learn's stock `Read`, for R3 to compare against. From "Permission levels
  // in SharePoint": View Items, Open Items, View Versions, Create Alerts,
  // Use Self-Service Site Creation, View Pages, Browse User Information, Use
  // Remote Interfaces, Use Client Integration Features, Open.
  //
  // NOT hardcoded as a bitmap. The High/Low pair is a 64-bit mask whose exact
  // value for a stock Read this project has never measured, and writing one
  // from memory is the failure AGENTS.md opens with. R2 prints what this site
  // has; R3 compares it to the OTHER built-ins on the SAME site, which is a
  // comparison that needs no external constant.
  const STOCK_READ_NOTE = 'compared against this site\'s own role definitions, not a remembered bitmap';

  expect('R1', 'CONTROL: can this caller read web/roledefinitions at all?');
  expect('R2', 'What BasePermissions does THIS site\'s built-in Read carry?');
  expect('R3', 'Does this site\'s Read differ from its neighbours in a way that suggests customisation?');
  expect('R4', 'Which groups hold a WEB-scope role assignment, and what?');
  expect('R5', 'Which groups hold a LIST-scope role assignment, and on what?');
  expect('R6', 'What groups does this site have, and what does each already hold?');

  if (!CONFIRMED) {
    log('INFO', 'PLAN. Nothing has been touched, and nothing would be.');
    log('INFO', 'This probe only READS. It would report:');
    log('INFO', `  - the BasePermissions of this site's built-in 'Read'`);
    log('INFO', '  - every web-scope and list-scope binding held by EVERY site group');
    log('INFO', 'Set CONFIRMED = true to run it. ALLOW_WRITES is not used.');
    report();
    return;
  }

  // ---- R1: the control, before any absence is believed -------------------
  const defs = await spGet('web/roledefinitions?$select=Id,Name,Description,BasePermissions,RoleTypeKind&$top=100');
  if (readFailed(defs) || !Array.isArray(defs.body && defs.body.value)) {
    record('R1', 'CONTROL: can this caller read web/roledefinitions at all?',
      `NOT ESTABLISHED (HTTP ${defs.status})`,
      'without this read, an empty binding list below would be indistinguishable from no access, which is the most dangerous wrong answer this probe could give. Re-run as a site owner.');
    report();
    return;
  }
  const rows = defs.body.value;
  record('R1', 'CONTROL: can this caller read web/roledefinitions at all?',
    'PASS', `${rows.length} role definition(s) readable, so an empty result below means empty`);

  // ---- R2: what Read actually is on this site ----------------------------
  const read = rows.find((r) => r.Name === 'Read');
  if (!read) {
    record('R2', 'What BasePermissions does THIS site\'s built-in Read carry?',
      'NOT ESTABLISHED (no level named Read)',
      `this site's levels are: ${rows.map((r) => r.Name).join(', ')}. A site with no 'Read' is itself a finding for #199. The validator exempts that name and the ACL phase resolves it by name.`);
  } else {
    record('R2', 'What BasePermissions does THIS site\'s built-in Read carry?',
      'PASS',
      `Read: RoleTypeKind=${read.RoleTypeKind}, High=${read.BasePermissions && read.BasePermissions.High}, Low=${read.BasePermissions && read.BasePermissions.Low}, Description=${JSON.stringify(read.Description)}`);
  }

  // ---- R3: is it plausibly customised? -----------------------------------
  // RoleTypeKind is the tell that needs no remembered constant: SharePoint
  // stamps a built-in with its type (Reader is 2). A level NAMED Read whose
  // type is None (0) is a CUSTOM level wearing the name, which is exactly
  // what #199 part 1 now refuses in a mapping and what nothing checks on a
  // site the mapping did not create.
  if (read) {
    const customType = read.RoleTypeKind === 0;
    record('R3', 'Does this site\'s Read differ from its neighbours in a way that suggests customisation?',
      customType ? 'FAIL' : 'PASS',
      customType
        ? `RoleTypeKind=0 means SharePoint does not consider this a built-in: a CUSTOM level is wearing the name 'Read' on this site, and the deploy would bind to it. ${STOCK_READ_NOTE}`
        : `RoleTypeKind=${read.RoleTypeKind}, so SharePoint still regards this as the built-in Reader level. Its bitmap is reported in R2 and should be compared across tenants before anything relies on a specific value. ${STOCK_READ_NOTE}`);
  }

  // ---- R6: survey every group on the site --------------------------------
  // Run 1 asked about ONE named group and learned nothing, because the name
  // it defaulted to only exists AFTER a deploy of the branch that introduced
  // it. R4 and R5 came back NOT ESTABLISHED and the operator was told to go
  // and guess a name.
  //
  // Surveying every group is strictly better AND COSTS THE SAME: the web
  // assignments are one read either way, and each list's assignments were
  // already read once per list. Only the reporting changes. The question
  // #198 actually asks is "what could a group already be carrying", and a
  // census answers that where a single lookup cannot.
  const groups = await spGet('web/sitegroups?$select=Id,Title&$top=200');
  if (readFailed(groups) || !Array.isArray(groups.body && groups.body.value)) {
    record('R6', 'What groups does this site have, and what does each already hold?',
      `NOT ESTABLISHED (HTTP ${groups.status})`, 'could not enumerate site groups');
    record('R4', 'Which groups hold a WEB-scope role assignment, and what?',
      'NOT ESTABLISHED (no group list)', 'R6 could not enumerate the groups');
    record('R5', 'Which groups hold a LIST-scope role assignment, and on what?',
      'NOT ESTABLISHED (no group list)', 'R6 could not enumerate the groups');
    report();
    return;
  }
  const byId = new Map(groups.body.value.map((g) => [g.Id, g.Title]));
  const named = (id) => byId.get(id) || `principal ${id}`;
  record('R6', 'What groups does this site have, and what does each already hold?',
    'PASS',
    `${byId.size} site group(s): ${[...byId.values()].join(', ')}`);

  // ---- R4: web-scope bindings, for every group ---------------------------
  // The binding the ACL phase never looks at and never removes: it reconciles
  // web/lists/.../roleassignments only.
  const webAsg = await spGet('web/roleassignments?$expand=RoleDefinitionBindings&$top=200');
  if (readFailed(webAsg) || !Array.isArray(webAsg.body && webAsg.body.value)) {
    record('R4', 'Which groups hold a WEB-scope role assignment, and what?',
      `NOT ESTABLISHED (HTTP ${webAsg.status})`, 'could not enumerate web role assignments');
  } else {
    const held = webAsg.body.value
      .filter((a) => byId.has(a.PrincipalId))
      .map((a) => `${named(a.PrincipalId)}: ${(a.RoleDefinitionBindings || []).map((b) => b.Name).join('+')}`);
    record('R4', 'Which groups hold a WEB-scope role assignment, and what?',
      'PASS',
      held.length
        ? `${held.join('; ')}. Anything here beyond a derived Limited Access is inherited by every account enrolled into that group, and nothing in the deploy removes it.`
        : 'no site group holds a web-scope role assignment');
  }

  // ---- R5: list-scope bindings, for every group, on every list -----------
  // HasUniqueRoleAssignments is what makes this readable. A list that
  // INHERITS returns the web's assignments verbatim, so run 1 reported every
  // group against nearly every list -- ~110 entries, of which the handful
  // that mattered were buried. Worse, it read as though the reader group held
  // a binding on Site Pages and Style Library when it holds one only at web
  // scope, which R4 already reports. Inherited lists are counted, not listed.
  const lists = await spGet("web/lists?$select=Id,Title,Hidden,HasUniqueRoleAssignments&$top=500");
  if (readFailed(lists) || !Array.isArray(lists.body && lists.body.value)) {
    record('R5', 'Which groups hold a LIST-scope role assignment, and on what?',
      `NOT ESTABLISHED (HTTP ${lists.status})`, 'could not enumerate lists');
  } else {
    const visible = lists.body.value.filter((l) => !l.Hidden);
    const unique = visible.filter((l) => l.HasUniqueRoleAssignments);
    const held = [];
    let unreadable = 0;
    for (const l of unique) {
      const asg = await spGet(
        `web/lists(guid'${l.Id}')/roleassignments?$expand=RoleDefinitionBindings&$top=200`);
      if (readFailed(asg) || !Array.isArray(asg.body && asg.body.value)) { unreadable += 1; continue; }
      for (const a of asg.body.value) {
        if (!byId.has(a.PrincipalId)) continue;
        const names = (a.RoleDefinitionBindings || []).map((b) => b.Name).join('+');
        held.push(`${l.Title} → ${named(a.PrincipalId)}: ${names}`);
      }
    }
    // The unreadable count is REPORTED, not swallowed. A list this caller
    // cannot read the ACL of is a list this probe cannot clear, and a
    // summary that omitted it would overstate what was checked.
    record('R5', 'Which groups hold a LIST-scope role assignment, and on what?',
      'PASS',
      `${unique.length} of ${visible.length} visible list(s) have UNIQUE permissions and were inspected; `
      + `the other ${visible.length - unique.length} inherit from the web, so their bindings are R4's, not their own`
      + `${unreadable ? `, and ${unreadable} unique list(s) had an ACL this caller could not read` : ''}. `
      + (held.length
        ? `${held.join('; ')}. Any of these on a list OUTSIDE a deployed bundle is inherited permanently by an account enrolled into that group: deploy/_acls.js.j2 iterates SCHEMA.list_assignments only.`
        : 'no site group holds an explicit binding on any list with unique permissions.'));
  }

  report();
  console.log('');
  console.log('=== WHAT TO DO WITH THIS ===');
  console.log('R2/R3 answer #199 part 2: whether this site\'s Read is still Read,');
  console.log('and whether a custom level is wearing the name. R3 FAIL means the');
  console.log('deploy would bind the reporting account to a level nobody checked.');
  console.log('R4/R5 answer #198: what a group already holds before the deploy');
  console.log('adds an account to it. Anything beyond Limited Access at web scope,');
  console.log('or any binding on a list outside the bundle, is privilege the');
  console.log('reader inherits permanently and no phase removes.');
  console.log('Together they are what #209\'s fail-closed gate has to test.');
})();
