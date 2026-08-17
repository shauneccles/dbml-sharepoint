/**
 * dbml-sharepoint PROBE: WHAT A CUSTOM PERMISSION LEVEL DOES ON THE WAY BACK
 *
 * WHY THIS EXISTS. `templates/deploy/_security_principals.js.j2` probes for a
 * role definition by name in `probeLevelExistence`, and when one already
 * exists `applyLevelDecision` MERGEs the declared `Description` and
 * `BasePermissions` onto it. It checks
 * `mergeResp.ok` and logs "declared permissions reconciled". It never reads
 * one field back, and it never asks where that level came from.
 *
 * A role definition is SITE-SCOPED. Its bitmap decides what every principal
 * holding it can do on every list that assigns it, including lists this tool
 * has never enumerated. Issue #224 is that hole. Issue #212 was the same
 * missing read-back for groups, and #209 was the same missing provenance
 * check for groups.
 *
 * `SP.RoleDefinition.Description` is a DIFFERENT SURFACE from
 * `SP.Group.Description` and from `SP.List.Description`. Both of those have
 * now been measured and neither result transfers here.
 * `analysis/permissions.py` says so explicitly beside the group ceiling:
 * "Custom permission LEVELS also carry a description and were not measured.
 * Do not assume this ceiling applies to them." This file measures it.
 *
 * WHAT TURNS ON THE ANSWER
 *
 *   - Whether a provenance marker can live on a role definition at all. If it
 *     cannot survive a write and a read, the adoption gate #224 needs cannot
 *     be built the way #223 built the group one, and the fix has to be
 *     something else.
 *   - Whether the deploy can gain a read-back verification of the bitmap it
 *     writes. R7 asks that directly. Today an HTTP 200 is the only evidence
 *     that a security-defining value was stored.
 *   - Whether an operator can be TOLD what adopting a level would affect.
 *     R9 asks whether the role assignments referencing a level can be
 *     enumerated. If they can, the gate can name the damage instead of
 *     guessing at it.
 *
 * SEPARATING WHAT IS DEPENDED ON FROM WHAT IS OBSERVED. R4 writes a long
 * description and REPORTS the length that comes back. It does not assert a
 * ceiling, because a control that asserts over the value it exists to
 * discover fails the moment the measurement starts working, and that looks
 * identical to a real refusal. R5 and R6 are the same: a value that does not
 * survive is an ANSWER recorded FAIL, not a broken run.
 *
 * THE BITMAP THIS PROBE SENDS IS AN INPUT, NOT A FINDING. The level is
 * created with a minimal viewer bitmap so that it is harmless if cleanup ever
 * failed and left it behind. Nothing here depends on that value being any
 * particular set of rights. R7 compares what comes back against what was
 * sent, whatever was sent. If SharePoint refuses the create outright, that is
 * recorded with its error rather than worked around, because a bitmap this
 * project cannot create is itself worth knowing.
 *
 * THE NEGATIVE CONTROL IS N1 AND IT IS NOT OPTIONAL. Every question below
 * reads a level back and compares. If a read of a level that does not exist
 * came back as a hit, every one of those comparisons would be meaningless.
 *
 * N1 USES THE $filter FORM DELIBERATELY, because that is what the deployer
 * uses, and the deployer's own comment says why:
 *
 *     getbyname returns HTTP 500 (not 404) for a missing role definition, so
 *     a getbyname probe cannot distinguish "absent" from a real failure.
 *
 * That claim has never been measured. R8 measures it, separately, and is
 * deliberately NOT the control: if the comment turns out to be wrong the
 * control still stands, because it never depended on the answer.
 *
 * WHAT THIS TOUCHES. One role definition, created and deleted by this probe
 * under a RUN-UNIQUE name it prints before doing anything. It assigns that
 * level to nobody, so it grants nothing to anyone at any point. It reads the
 * web's existing role assignments in R9 and writes to none of them. It
 * touches no list, no item, no group, and no level it did not create, and if
 * the name is somehow already taken it stops without writing or deleting.
 * It needs ManagePermissions, so a site owner.
 *
 * HOW TO RUN
 *   1. Paste it once: it prints the web and stops. Set CONFIRMED = true.
 *   2. Open that site's classic settings page, /_layouts/15/settings.aspx,
 *      signed in as a Site Owner. The site guard needs _spPageContextInfo.
 *   3. F12 -> Console -> type `allow pasting` if the browser objects ->
 *      paste this whole file -> Enter.
 *   4. Set ALLOW_WRITES = true and paste again to actually run it.
 *   5. Copy the whole RESULTS block back verbatim.
 *
 * RUN 1, 2026-08-14, revision 709c2549, one Microsoft 365 group-connected
 * Team Site. Ten questions, ten answered. Both controls held: N1 read HTTP
 * 200 with zero rows for an absent name, and the level was created, read and
 * deleted under its run-unique name.
 *
 *   THE CEILING IS 512 AND IT REFUSES RATHER THAN TRUNCATES. R4 sent 1018
 *   characters and got HTTP 500:
 *
 *       "The parameter Description cannot be bigger than 512 characters."
 *
 *   The same number as a site group, and a refusal in the same style, but
 *   NOT the same message: the group surface says "cannot be null or bigger
 *   than 512 characters" and this one omits the null clause. Two properties
 *   that happen to agree today, on different types, measured separately.
 *   R4 is recorded FAIL because 1018 characters did not come back, which is
 *   the ANSWER the question was asked to get, not a broken run.
 *
 *   EVERYTHING A MARKER NEEDS WORKS.
 *     R1  byte-identical through the CREATE path
 *     R2  byte-identical through the MERGE path
 *     R3  MERGE is PARTIAL: the Description survived a request that did not
 *         carry it. The non-vacuity was NOT established in run 1, see the
 *         correction below, and was settled in run 2
 *     R5  an empty description is accepted and reads back ""
 *     R6  the marker shape lists and groups already use survives intact
 *
 *   R7 THE BITMAP READS BACK AS WRITTEN. High=0 Low=196609 out, the same
 *   back. So the deploy CAN verify the most security-defining value it
 *   writes, which today it sends and never reads. That is the #212 gap for
 *   role definitions and it is now closable.
 *
 *   R8 CONFIRMS THE DEPLOYER COMMENT. `getbyname` really does answer HTTP
 *   500, not 404, for a level that is not there. The $filter existence probe
 *   in `_security_principals.js.j2` is necessary and the comment above it is
 *   now measured rather than asserted.
 *
 *   R9 WEB-SCOPE ASSIGNMENTS ARE ENUMERABLE. Ten role assignments read, all
 *   ten with an expandable RoleDefinitionBindings collection, and no
 *   continuation link at a page size of 200. SEE THE CORRECTION BELOW: run 1
 *   checked that the collection was an array, NOT that its entries expose an
 *   Id or Name a gate could match a level against. PER-LIST usage is a
 *   different and much more expensive question that stays unanswered.
 *
 * WHAT RUN 1 IMMEDIATELY PRODUCED, beyond the questions asked. Nothing in
 * the build guards this ceiling. A mapping declaring a level description
 * over 512 characters builds clean and fails at phase 1.2, after permission
 * levels have begun being written and before any list exists. The longest
 * description in the shipped catalogue is 166 characters, so no family is at
 * risk today, and that is luck rather than a guard. Filed as #227.
 *
 * CORRECTION TO RUN 1, made 2026-08-14 after review, before any of this was
 * built on. R3's recorded evidence said the omitted-field MERGE "did change
 * Order in the same call". THE PROBE NEVER READ ORDER BACK. It did not even
 * select it. That sentence described the request rather than reporting a
 * measurement, so run 1 established the weaker claim only:
 *
 *     MEASURED: a MERGE that omitted Description was accepted, and the
 *     Description was unchanged afterwards.
 *     NOT MEASURED: that the MERGE did anything at all. A server that
 *     ignored the entire request produces the same observation.
 *
 * The code now reads Order BEFORE the MERGE, derives a control value that
 * differs from what it actually observed rather than from what the create
 * request sent, reads Order back, and records NOT ESTABLISHED when the
 * control field did not move. Deriving the control from the observed value
 * matters for the same reason the first correction did: had the tenant
 * normalised Order at create time to the number the MERGE was about to
 * write, an ignored MERGE would have read back as the target and the control
 * would have passed while proving nothing.
 *
 * SECOND CORRECTION, same review. R9 recorded PASS after checking that each
 * assignment's RoleDefinitionBindings was an ARRAY. A gate does not need an
 * array, it needs an Id or a Name to match a level against, and run 1 never
 * checked for one. Run 1 therefore established that the collection is
 * present and expandable, and NOT that it is usable. The code now requires
 * every row to expose an Id or Name, follows every continuation page, and
 * records NOT ESTABLISHED on a truncated enumeration, because the
 * assignments a partial read misses are exactly the usage a gate would fail
 * to report.
 *
 * RUN 2, 2026-08-14, revision 4dd7de1c, same site, and it closes both
 * corrections. Ten questions, ten answered, every run-1 result reproduced.
 *
 *   R3 IS NOW MEASURED, not described. Order read as 100 BEFORE the MERGE,
 *   the control value 101 was derived from that observation, and Order read
 *   back as 101 afterwards while the Description was untouched. So the MERGE
 *   demonstrably did something and the Description survived it anyway.
 *   Partial MERGE holds on role definitions.
 *
 *   R9 IS NOW ESTABLISHED at the strength a gate needs. Ten assignments over
 *   ONE page with no continuation left, and EVERY one exposed a level Id or
 *   Name. Not merely an array this time.
 *
 *   CLEANUP VERIFIED. "Deleted, and confirmed absent on read-back", so the
 *   DELETE was not taken on the strength of its own status code.
 *
 * WHAT IS SETTLED AFTER TWO RUNS. A provenance marker can live on a role
 * definition: byte-exact through both write paths, marker shape survives,
 * empty accepted, partial MERGE confirmed. The bitmap can be verified. A
 * refusal can name web-scope usage. The ceiling is 512 and it refuses rather
 * than truncates. PER-LIST usage remains the one open question, deliberately
 * not asked here.
 *
 * OPEN QUESTIONS FOR A FUTURE RUN
 *
 *   - CASE SENSITIVITY OF `$filter=Name eq '...'`. Not documented on
 *     Microsoft Learn and not one of the ten questions above. The existence
 *     probe in `_security_principals.js.j2` uses this filter to decide
 *     whether a declared level already exists. If the match is
 *     case-sensitive, a site level named 'schema manager' reads as absent
 *     against a declared 'Schema Manager', the create path runs, and the
 *     #224 adoption gate never sees the existing level. That is NOT a gate
 *     fail-open: nothing is written to the existing level, which is the same
 *     behaviour as before #224. The residual is real: the create either
 *     collides with the existing level and errors, or two case-variant
 *     levels end up coexisting, and `_acls.js.j2`'s `resolveRoleDefId`, which
 *     resolves a level by name through its own `getbyname` call, then binds
 *     whichever one SharePoint's name resolution picks. No result is
 *     recorded for this yet; it is a question, not a finding.
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
  log('INFO', 'probe revision 11905545. Quote this when reporting results.');

  // Run-unique, and that is a safety property rather than tidiness. A fixed
  // name plus a pre-emptive delete destroys somebody else's level on the one
  // site where the name is already taken, and a role definition in use is
  // exactly the object this probe exists to protect.
  const RUN = `${Date.now().toString(36)}`.slice(-6);
  const LEVEL = `dbmlsp Probe Level ${RUN}`;

  // Mirrors the marker shape `analysis/group_description.py` composes, because
  // R6 asks whether THAT string could live on a role definition. A marker that
  // cannot survive the round trip cannot be the basis of #224's gate.
  const MARKER = 'Provisioned by dbml-sharepoint from probe/Level.';

  const PLAIN = 'dbmlsp probe plain level description';
  const MERGED = 'dbmlsp probe merged level description';
  const LONG = `dbmlsp probe long ${'x'.repeat(1000)}`;

  // A minimal viewer bitmap: harmless if cleanup ever failed and left the
  // level behind. See the preamble; nothing here depends on the value.
  const BASE_LOW = '196609';
  const BASE_HIGH = '0';

  expect('N1', 'CONTROL: does the $filter probe report a level that does not exist as absent?');
  expect('R1', 'Does a plain Description written at CREATE read back byte-identical?');
  expect('R2', 'Does a Description written by MERGE read back byte-identical?');
  expect('R3', 'Does a MERGE that OMITS Description preserve the previous one?');
  expect('R4', 'What length comes back when 1000+ characters go out?');
  expect('R5', 'Is an EMPTY description accepted?');
  expect('R6', 'Does a provenance marker of the shape lists and groups use survive?');
  expect('R7', 'Does BasePermissions read back as written?');
  expect('R8', 'What does getbyname actually return for an ABSENT level?');
  expect('R9', 'Can the role assignments referencing a level be enumerated?');

  if (!CONFIRMED || !ALLOW_WRITES) {
    log('INFO', 'PLAN. Nothing has been touched.');
    log('INFO', `This probe would create the permission level '${LEVEL}', write and`);
    log('INFO', 'read back its Description several ways, read its bitmap back, read');
    log('INFO', "the web's existing role assignments, then delete the level.");
    log('INFO', 'It assigns the level to nobody, so it grants nothing to anyone.');
    log('INFO', 'Set CONFIRMED = true and ALLOW_WRITES = true to run it.');
    report();
    return;
  }

  // Mirrors the deployer's own request framing exactly, so a failure here
  // means "the description did not survive" rather than "the probe asked
  // differently from the deployer".
  const VERBOSE = { 'Content-Type': 'application/json;odata=verbose' };

  const filterLevel = async (name) =>
    spGet(`web/roledefinitions?$select=Id,Name,Description&$filter=Name eq '${encodeURIComponent(name)}'`);

  // Order is selected because R3's non-vacuity depends on reading it back,
  // not on the request having carried it.
  const readLevelById = async (id) =>
    spGet(`web/roledefinitions(${id})?$select=Id,Name,Description,BasePermissions,Order`);

  let digest = await getDigest();
  log('INFO', `This run uses the level name '${LEVEL}'.`);

  // ---- N1: the negative control, and the existence check, in one read ----
  // Nothing is deleted first. If the name is somehow taken, that is a level
  // this probe did not create and it stops rather than touching it.
  const absent = await filterLevel(LEVEL);
  if (readFailed(absent)) {
    record('N1', 'CONTROL: does the $filter probe report a level that does not exist as absent?',
      `NOT ESTABLISHED (HTTP ${absent.status})`,
      'the control needs a clean read for a name that is not there; nothing below would mean anything');
    report();
    return;
  }
  const absentRows = (absent.body && (absent.body.value || absent.body.results)) || [];
  if (absentRows.length !== 0) {
    record('N1', 'CONTROL: does the $filter probe report a level that does not exist as absent?',
      'NOT ESTABLISHED (name already taken)',
      `'${LEVEL}' already exists on this site. Nothing was written or deleted. Re-run for a different name, or delete that level by hand if a previous run left it.`);
    report();
    return;
  }
  record('N1', 'CONTROL: does the $filter probe report a level that does not exist as absent?',
    'PASS',
    `HTTP ${absent.status} with zero rows for a name that is not there, so a hit below means something`);

  // ---- R8: what getbyname does for an absent level ----------------------
  // Asked BEFORE the level exists, and deliberately not the control. The
  // deployer's comment claims HTTP 500 here. If that is wrong the comment
  // should say so; if it is right, the $filter form stays necessary.
  const byNameAbsent = await spGet(`web/roledefinitions/getbyname('${encodeURIComponent(LEVEL)}')?$select=Id`);
  record('R8', 'What does getbyname actually return for an ABSENT level?',
    byNameAbsent.status === 500 ? 'PASS (500, as the deployer comment claims)'
      : byNameAbsent.status === 404 ? 'FAIL (404, so the comment is wrong and $filter may be unnecessary)'
        : `NOT ESTABLISHED (HTTP ${byNameAbsent.status})`,
    `HTTP ${byNameAbsent.status} for a level that is not there`);

  // ---- R1 and R7: create with a plain description and a known bitmap ----
  digest = await getDigest();
  const created = await spPost('web/roledefinitions', {
    __metadata: { type: 'SP.RoleDefinition' },
    Name: LEVEL,
    Description: PLAIN,
    BasePermissions: {
      __metadata: { type: 'SP.BasePermissions' },
      High: BASE_HIGH,
      Low: BASE_LOW,
    },
    Order: 100,
  }, digest, VERBOSE);
  if (!created.ok) {
    record('BOOT', 'create the probe permission level',
      `NOT ESTABLISHED (HTTP ${created.status})`, created.text.slice(0, 300));
    report();
    return;
  }
  let levelId = created.body && (created.body.Id || (created.body.d && created.body.d.Id));
  if (!levelId) {
    // The create SUCCEEDED, so something exists on this site under a name
    // this probe chose. Resolve it rather than walking away: returning here
    // without an Id is how a probe that promises to clean up after itself
    // leaves a site-scoped object behind.
    const resolved = await filterLevel(LEVEL);
    const resolvedRows = (resolved.body && (resolved.body.value || resolved.body.results)) || [];
    levelId = resolvedRows.length === 1 ? resolvedRows[0].Id : null;
    if (!levelId) {
      record('BOOT', 'create the probe permission level',
        'NOT ESTABLISHED (no Id returned, and the name could not be resolved)',
        `the create returned HTTP ${created.status} without an Id and '${LEVEL}' could not be resolved by name. It may exist. Delete it by hand in Site permissions > Permission Levels and report this.`);
      report();
      return;
    }
    log('INFO', `Create returned no Id; resolved '${LEVEL}' to Id ${levelId} so it can still be cleaned up.`);
  } else {
    log('INFO', `Created level Id ${levelId}.`);
  }

  // EVERYTHING FROM HERE RUNS INSIDE try/finally. A rejected fetch, a
  // getDigest that throws on a throttle, or a JSON parse failure would
  // otherwise skip the cleanup below and leave a site-scoped permission level
  // behind, under a transcript that never mentioned it. The preamble promises
  // this probe deletes what it creates, and a promise that holds only on the
  // happy path is not one.
  try {
  const afterCreate = await readLevelById(levelId);
  if (readFailed(afterCreate)) {
    record('R1', 'Does a plain Description written at CREATE read back byte-identical?',
      `NOT ESTABLISHED (HTTP ${afterCreate.status})`, 'the level was created but could not be read back');
  } else {
    const got = afterCreate.body.Description;
    record('R1', 'Does a plain Description written at CREATE read back byte-identical?',
      got === PLAIN ? 'PASS' : 'FAIL',
      `sent ${JSON.stringify(PLAIN)}, read ${JSON.stringify(got)}`);
  }

  if (!readFailed(afterCreate)) {
    const bp = afterCreate.body.BasePermissions || {};
    const sameLow = String(bp.Low) === String(BASE_LOW);
    const sameHigh = String(bp.High) === String(BASE_HIGH);
    record('R7', 'Does BasePermissions read back as written?',
      sameLow && sameHigh ? 'PASS' : 'FAIL',
      `sent High=${BASE_HIGH} Low=${BASE_LOW}, read High=${bp.High} Low=${bp.Low}`);
  }

  // Every remaining Description question is the same shape, so it is one
  // function rather than four copies of it.
  const mergeDescription = async (payload, extra = {}) => {
    const d = await getDigest();
    return spPost(
      `web/roledefinitions(${levelId})`,
      { __metadata: { type: 'SP.RoleDefinition' }, ...payload },
      d,
      { ...VERBOSE, 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*', ...extra },
    );
  };

  const askDescription = async (id, question, sent, payload) => {
    const res = await mergeDescription(payload);
    if (!res.ok) {
      record(id, question,
        isRefusal(res.status) ? 'FAIL' : `NOT ESTABLISHED (HTTP ${res.status})`,
        `HTTP ${res.status}: ${res.text.slice(0, 200)}`);
      return null;
    }
    const back = await readLevelById(levelId);
    if (readFailed(back)) {
      record(id, question, `NOT ESTABLISHED (HTTP ${back.status})`,
        'the MERGE succeeded but the level could not be read back');
      return null;
    }
    const got = back.body.Description;
    if (sent !== null) {
      record(id, question, got === sent ? 'PASS' : 'FAIL',
        `sent ${JSON.stringify(sent)}, read ${JSON.stringify(got)}`);
    }
    return got;
  };

  // ---- R2: MERGE path round trip ----------------------------------------
  await askDescription('R2', 'Does a Description written by MERGE read back byte-identical?',
    MERGED, { Description: MERGED });

  // ---- R3: is MERGE partial? --------------------------------------------
  // Non-vacuous ONLY IF the request demonstrably did something. It carries a
  // new Order, and that Order is READ BACK below. Run 1 claimed the
  // non-vacuity in its evidence string without ever checking it, which made
  // a surviving Description equally explainable by the server having ignored
  // the whole request. A control that asserts its own premise is not a
  // control, so the premise is measured here instead of described.
  // The control value is DERIVED from what the level actually reads as, not
  // from what the create request sent. Assuming the pre-state is the same
  // mistake one level up: if the tenant normalised Order at create time to
  // whatever we were about to write, an ignored MERGE would read back as the
  // target and the control would pass while proving nothing.
  const beforeOmit = await readLevelById(levelId);
  if (readFailed(beforeOmit)) {
    // Recorded and stepped over, never returned from: the remaining questions
    // do not depend on R3, and an early return here would be a second way to
    // skip the cleanup this probe promises.
    record('R3', 'Does a MERGE that OMITS Description preserve the previous one?',
      `NOT ESTABLISHED (HTTP ${beforeOmit.status})`,
      'the level could not be read before the MERGE, so no control value could be chosen');
  }
  const ORDER_BEFORE = readFailed(beforeOmit) ? NaN : Number(beforeOmit.body.Order);
  const ORDER_AFTER = ORDER_BEFORE === 101 ? 102 : 101;
  const omitRes = Number.isNaN(ORDER_BEFORE)
    ? { ok: false, status: 0, text: 'skipped: no pre-MERGE Order could be read' }
    : await mergeDescription({ Order: ORDER_AFTER });
  if (!omitRes.ok) {
    record('R3', 'Does a MERGE that OMITS Description preserve the previous one?',
      isRefusal(omitRes.status) ? 'FAIL' : `NOT ESTABLISHED (HTTP ${omitRes.status})`,
      `HTTP ${omitRes.status}: ${omitRes.text.slice(0, 200)}`);
  } else {
    const backOmit = await readLevelById(levelId);
    if (readFailed(backOmit)) {
      record('R3', 'Does a MERGE that OMITS Description preserve the previous one?',
        `NOT ESTABLISHED (HTTP ${backOmit.status})`,
        'the MERGE succeeded but the level could not be read back');
    } else {
      const gotDesc = backOmit.body.Description;
      const gotOrder = Number(backOmit.body.Order);
      // Both halves matter. Reaching the target says the request was applied;
      // differing from the observed pre-value says the request is why.
      const orderMoved = gotOrder === ORDER_AFTER && gotOrder !== ORDER_BEFORE;
      if (!orderMoved) {
        // The premise failed, so the question is UNANSWERED rather than
        // answered yes. A preserved Description here proves nothing.
        record('R3', 'Does a MERGE that OMITS Description preserve the previous one?',
          'NOT ESTABLISHED (the control field did not move)',
          `Order read ${ORDER_BEFORE} before the MERGE, the request carried ${ORDER_AFTER}, and it read back ${gotOrder}. This MERGE cannot be shown to have done anything, so a surviving Description would be equally explained by the whole request being ignored. Description read ${JSON.stringify(gotDesc)}`);
      } else {
        record('R3', 'Does a MERGE that OMITS Description preserve the previous one?',
          gotDesc === MERGED ? 'PASS' : 'FAIL',
          gotDesc === MERGED
            ? `Description survived a MERGE that did not carry it, and the same call DID move Order from an OBSERVED ${ORDER_BEFORE} to ${gotOrder}, so the request was not ignored. Description read ${JSON.stringify(gotDesc)}`
            : `Description did NOT survive a partial MERGE; expected ${JSON.stringify(MERGED)}, read ${JSON.stringify(gotDesc)}. Order moved from an observed ${ORDER_BEFORE} to ${gotOrder}, so the request was applied`);
      }
    }
  }

  // ---- R4: length, REPORTED not asserted --------------------------------
  const longBack = await askDescription('R4', 'What length comes back when 1000+ characters go out?',
    null, { Description: LONG });
  if (longBack !== null) {
    record('R4', 'What length comes back when 1000+ characters go out?',
      longBack.length === LONG.length ? `PASS (${longBack.length} of ${LONG.length})` : `SHORT (${longBack.length} of ${LONG.length})`,
      longBack.length === LONG.length
        ? `sent ${LONG.length} characters and read the same number back, so the accepted length is at least ${LONG.length}`
        : `sent ${LONG.length} characters and read ${longBack.length} back, so this surface truncates rather than refuses`);
  }

  // ---- R6: the marker shape ---------------------------------------------
  await askDescription('R6', 'Does a provenance marker of the shape lists and groups use survive?',
    MARKER, { Description: MARKER });

  // ---- R5: empty ---------------------------------------------------------
  await askDescription('R5', 'Is an EMPTY description accepted?', '', { Description: '' });

  // ---- R9: what would an adoption affect? --------------------------------
  // Read-only. Answers whether #224's gate can NAME the damage rather than
  // warn in the abstract. Web scope only: a per-list sweep is a different and
  // much more expensive question, and this one has to be answered first.
  // EVERY page, or the question is unanswered. A first page carrying one
  // identifiable binding would otherwise certify that web-scope usage is
  // reportable while later assignments were never looked at, which is the
  // same fail-open a truncated membership enumeration would be.
  let assignments = await spGet(
    'web/roleassignments?$select=PrincipalId&$expand=RoleDefinitionBindings&$top=200',
  );
  const rows = [];
  let truncated = false;
  let pages = 0;
  if (!readFailed(assignments)) {
    let body = assignments.body;
    for (;;) {
      pages += 1;
      rows.push(...((body && (body.value || body.results)) || []));
      const next = body && (body['odata.nextLink'] || body.__next);
      if (!next) break;
      // A continuation link is absolute, so it does not go through spGet.
      const res = await fetch(next, { headers: { Accept: 'application/json;odata=nometadata' } });
      if (!res.ok) { truncated = true; break; }
      body = await res.json().catch(() => null);
      if (!body) { truncated = true; break; }
      if (pages > 50) { truncated = true; break; }
    }
  }
  if (readFailed(assignments)) {
    record('R9', 'Can the role assignments referencing a level be enumerated?',
      `NOT ESTABLISHED (HTTP ${assignments.status})`,
      'the web role assignments could not be read, so the gate cannot report what an adoption would affect');
  } else if (truncated) {
    record('R9', 'Can the role assignments referencing a level be enumerated?',
      'NOT ESTABLISHED (enumeration truncated)',
      `read ${rows.length} row(s) over ${pages} page(s) and then could not follow the continuation. A partial enumeration cannot establish that usage is reportable, because the assignments it did not reach are exactly the ones a gate would miss.`);
  } else {
    // The gate needs the BINDINGS, not the assignment rows. A row whose
    // RoleDefinitionBindings cannot be parsed tells it nothing about which
    // level is in use, so counting rows alone would certify the capability
    // on evidence that does not carry it.
    const bindingsOf = (r) => {
      const b = r.RoleDefinitionBindings;
      if (Array.isArray(b)) return b;
      if (b && Array.isArray(b.results)) return b.results;
      return null;
    };
    const parsed = rows.map(bindingsOf).filter((b) => b !== null);
    // Does a binding actually expose the level identity the gate would match on?
    const identified = parsed.filter(
      (b) => b.some((d) => d && (d.Id !== undefined || d.Name !== undefined)),
    ).length;
    if (rows.length === 0) {
      record('R9', 'Can the role assignments referencing a level be enumerated?',
        'NOT ESTABLISHED (no role assignments on this web)',
        'this web reported no role assignments at all, which is unusual; the question is unanswered rather than answered no');
    } else if (identified !== rows.length) {
      // EVERY row, not merely one. A gate that can identify the level behind
      // some assignments and not others reports an incomplete list of what an
      // adoption would affect, and an incomplete list read as a complete one
      // is worse than no list.
      record('R9', 'Can the role assignments referencing a level be enumerated?',
        'NOT ESTABLISHED (identity coverage incomplete)',
        `read ${rows.length} row(s) over ${pages} page(s); ${parsed.length} had a parseable RoleDefinitionBindings collection and ${identified} exposed a level Id or Name. A gate needs all of them, because the rows it cannot identify are the usage it would fail to report.`);
    } else {
      record('R9', 'Can the role assignments referencing a level be enumerated?',
        'PASS',
        `read ${rows.length} web-scope role assignment(s) over ${pages} page(s) with no continuation left, and every one exposed a level Id or Name a gate could match on. Web-scope usage is therefore reportable. PER-LIST usage is a separate and much more expensive question this run does not answer.`);
    }
  }

  } catch (err) {
    // The measurement that threw is already NOT ESTABLISHED, because every
    // question was registered up front. Recorded here so the transcript says
    // the run stopped early rather than leaving the reader to infer it.
    log('FAIL', `The probe stopped after the level was created: ${(err && err.message) || err}. Cleanup still runs below.`);
  } finally {

  // ---- Clean up ----------------------------------------------------------
  // By Id, never by name: the Id is the one this run created and nothing else.
  //
  // Wrapped in its own try/catch because this whole block is a `finally`. An
  // exception thrown from here would replace whatever sent the run here in
  // the first place, so a throttled digest during cleanup would hide the
  // original failure and still not delete anything.
  try {
  digest = await getDigest();
  const removed = await spPost(
    `web/roledefinitions(${levelId})`, {}, digest,
    { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' },
  );
  // Cleanup is a write, and this project does not accept an HTTP 200 as
  // evidence that a write took. A DELETE reported as succeeding while the
  // level is still there would leave a site-scoped object behind under a
  // transcript claiming the site was left clean, which is the worst of both.
  if (removed.ok) {
    const gone = await filterLevel(LEVEL);
    const goneRows = (gone.body && (gone.body.value || gone.body.results)) || [];
    if (!readFailed(gone) && goneRows.length === 0) {
      log('OK', `Deleted '${LEVEL}', and confirmed absent on read-back.`);
    } else {
      log('FAIL', `DELETE reported success for '${LEVEL}' but it is STILL PRESENT on read-back (HTTP ${gone.status}, ${goneRows.length} row(s)). Remove it by hand in Site permissions > Permission Levels, and report this: a DELETE that reports success without deleting is itself a finding.`);
    }
  } else {
    log('FAIL', `Could not delete '${LEVEL}' (HTTP ${removed.status}). Remove it by hand in Site permissions > Permission Levels.`);
  }
  } catch (cleanupErr) {
    log('FAIL', `Cleanup of '${LEVEL}' threw: ${(cleanupErr && cleanupErr.message) || cleanupErr}. The level Id was ${levelId}. Remove it by hand in Site permissions > Permission Levels.`);
  }
  }

  report();
  console.log('');
  console.log('=== WHAT TO DO WITH THIS ===');
  console.log('R1/R2/R6 decide whether a provenance marker can live on a role');
  console.log('definition at all, and so whether #224 can be gated the way #223');
  console.log('gated groups.');
  console.log('R3 decides whether the deploy can write one field without');
  console.log('disturbing another, as it already relies on for groups.');
  console.log('R4 reports the length that came back. It does NOT assert a ceiling.');
  console.log('R7 decides whether the deploy can verify the bitmap it writes,');
  console.log('which today it sends and never reads back.');
  console.log('R8 either confirms or retires the deployer comment claiming');
  console.log('getbyname answers 500 for an absent level.');
  console.log('R9 decides whether a refusal can NAME what adopting a level would');
  console.log('affect, instead of warning in the abstract.');
})();
