/**
 * dbml-sharepoint PROBE: WHAT A DOCUMENT LIBRARY DOES WITH THE STANDARD
 *
 * QUESTION: can a `kind: DocumentLibrary` carry declared views, a form
 * header and demo rows, and if not, what should it carry instead?
 *
 * WHY: `policy-library` is the LAST entry on the family sweep's
 * NOT_YET_UPLIFTED roster. The family-standard spec said that set must
 * reach empty; it reached one, and this is the one. Its PolicyRegister half
 * is uplifted in full. Its PolicyDocuments half is a document library, and
 * three parts of the standard do not describe one.
 *
 * Three of the four blockers are already settled WITHOUT a tenant, by
 * building them (see the header comment in
 * templates/policy-library/20-configure/mapping.yaml):
 *
 *   - a library view naming FileLeafRef is a BUILD error, because the tool
 *     accepts DBML columns plus ID/Created/Modified/Author/Editor and
 *     nothing else;
 *   - a form body naming it fails the same way;
 *   - the standard's header title line is `[$Title]`, and on a library
 *     Title is a separate, normally empty field.
 *
 * The fourth is not a build error and is worse, and it is why this probe
 * exists: demo_items on a library generate a POST to /items, which asks
 * SharePoint to create an item WITH NO FILE BEHIND IT. That has never been
 * run against a tenant. It is currently a build error taken on suspicion
 * (a reasonable place to stand, but not an answered question).
 *
 * WHAT IT ASKS
 *   LN   NEGATIVE CONTROL: is a POST to /items naming a column that does
 *        not exist REFUSED? Establishes this probe can see a failed item
 *        write at all. Without it, L2 refusing proves nothing.
 *   L1   does a document library create at all (BaseTemplate 101)?
 *   L2   THE ONE THAT MATTERS. Does POST /items on a library succeed?
 *        Refused outright would be the happy answer: demo data simply
 *        cannot be generated and the build error stands on evidence.
 *   L3   ...and if it SUCCEEDS, what was created? FileSystemObjectType,
 *        FileRef, FileLeafRef and Title are read back. A "ghost" (an item
 *        with no file) is the outcome the build error was guessing at,
 *        and this is what proves or disproves it.
 *   L4   ...and does that ghost APPEAR to a user? An invisible row and a
 *        visible broken one are very different problems: one is untidy,
 *        the other puts a document-shaped nothing in front of staff.
 *   L5   after uploading a REAL file, is Title actually empty? This is the
 *        assumption behind refusing the `[$Title]` header line. If
 *        SharePoint populates it after all, the header could ship as-is.
 *   L6   can a library VIEW carry FileLeafRef through REST? The tool
 *        refuses it; this establishes whether that is a TOOL limit that
 *        could be lifted, or a platform one that cannot.
 *   L7   does a library content type accept a ClientFormCustomFormatter
 *        whose title line reads [$FileLeafRef] instead of [$Title]?
 *        STORAGE ONLY: whether it renders needs eyes, and the checklist
 *        at the end asks for them.
 *
 * L7 HAS NOT BEEN RE-RUN SINCE 2026-07-29. The run on that date took the
 * first content type the collection happened to return, and a library
 * carries a Folder content type as well as a Document one; which of the two
 * accepted the formatter was never established. The selection is now by id,
 * the way the deploy does it, but nobody has pasted the corrected script
 * into a tenant. Treat L7's recorded STORED as a result about an unnamed
 * content type until it is run again. Nothing shipped rests on it: the
 * refusal of `kind: DocumentLibrary` names L2 (a POST to /items is refused)
 * and L5 (Title is null after an upload), plus the fact that this tool has
 * no upload step, none of which touch the content type L7 wrote to.
 *
 * READ LN FIRST. If a nonsense item write is ACCEPTED, this probe cannot
 * distinguish refusal from success and L2 is unproven rather than answered.
 *
 * HOW TO RUN
 *   1. Open a site you own, at /_layouts/15/settings.aspx.
 *   2. F12 -> Console -> paste -> Enter. It prints its plan and stops.
 *   3. Edit CONFIRMED and ALLOW_WRITES to true, paste again.
 *   4. Copy the RESULTS block back, and answer the short eyes-on checklist
 *      it prints.
 *
 * WHEN FINISHED: delete the library it created.
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

  const LIB = 'dbmlsp Probe DocLib';
  const listPath = `web/lists/getbytitle('${LIB}')`;

  if (!CONFIRMED) {
    log('INFO', `Would create a DOCUMENT LIBRARY '${LIB}' on ${WEB}, then try to`);
    log('INFO', 'POST an item to it with no file behind it, read back what that');
    log('INFO', 'produced, upload one real file to see whether Title is populated,');
    log('INFO', 'add FileLeafRef to a view, and store a form header referencing it.');
    if (CLEANUP) {
      log('INFO', `CLEANUP is ON: '${LIB}' would be RECYCLED first, with its contents.`);
    } else {
      log('INFO', `CLEANUP is off: an existing '${LIB}' would be reused, and rows`);
      log('INFO', 'from a previous run would answer this run\'s questions.');
      log('INFO', 'Set CLEANUP = true for a clean run.');
    }
    log('INFO', 'Nothing has been written. Set CONFIRMED and ALLOW_WRITES to true.');
    return;
  }
  if (!ALLOW_WRITES) {
    log('INFO', 'CONFIRMED, but ALLOW_WRITES is false and this probe must write.');
    log('INFO', 'Set ALLOW_WRITES = true to proceed. Stopping.');
    return;
  }

  expect('LN', 'NEGATIVE CONTROL: an item POST naming a missing column is refused');
  expect('L1', 'A document library is created (BaseTemplate 101)');
  expect('L2', 'POST /items on a document library');
  expect('L3', 'What a fileless item POST actually created');
  expect('L4', 'Whether that item is visible to a user');
  expect('L5', 'Is Title populated after a real file upload');
  expect('L6', 'Can a library view carry FileLeafRef through REST');
  expect('L7', 'A library content type accepts a header referencing [$FileLeafRef]');

  await resetList(LIB);
  let digest = await getDigest();

  // ---- L1: the library ------------------------------------------------
  const existing = await spGet(listPath);
  if (existing.ok) {
    record('L1', 'A document library is created (BaseTemplate 101)', 'ALREADY PRESENT',
           'reusing an existing library. Set CLEANUP = true for a clean answer');
  } else {
    const made = await spPost('web/lists', {
      Title: LIB,
      BaseTemplate: 101,
      Description: 'dbml-sharepoint probe library. Safe to delete.',
    }, digest);
    record('L1', 'A document library is created (BaseTemplate 101)',
           made.ok ? 'PASS' : 'FAIL',
           made.ok ? `created '${LIB}'` : `HTTP ${made.status}: ${made.text.slice(0, 300)}`);
    if (!made.ok) return report();
  }

  // ---- LN: NEGATIVE CONTROL -------------------------------------------
  digest = await getDigest();
  const junk = await spPost(`${listPath}/items`,
                            { NoSuchColumnAtAll: 'x' }, digest);
  record('LN', 'NEGATIVE CONTROL: an item POST naming a missing column is refused',
         junk.ok ? 'FAIL' : isRefusal(junk.status) ? 'PASS' : 'NOT ESTABLISHED',
         junk.ok
           ? 'an item POST naming a column that does not exist was ACCEPTED. This '
             + 'probe cannot tell a refused item write from a successful one, so L2 '
             + 'is unproven whichever way it goes'
           : isRefusal(junk.status)
             ? `refused with HTTP ${junk.status}: ${junk.text.slice(0, 260)}`
             : `the request failed with HTTP ${junk.status}, which is not the server `
               + 'refusing the write. L2 is unproven rather than answered: '
               + junk.text.slice(0, 200));

  // ---- L2 / L3 / L4: the fileless item --------------------------------
  // This is the exact shape demogen would emit for a library: a plain
  // /items POST carrying a Title and nothing else.
  digest = await getDigest();
  const ghost = await spPost(`${listPath}/items`,
                             { Title: '[DEMO] probe fileless item' }, digest);
  if (!ghost.ok) {
    // THE row the shipped refusal rests on, so it says which kind of "no"
    // it got. The recorded 2026-07-29 answer was a 500 carrying
    // "To add an item to a document library, use SPFileCollection.Add()",
    // which is content-specific, and a 500 on its own is not: a throttle or
    // an outage lands in the same branch. What makes it evidence is the
    // MESSAGE, so the message is what the verdict turns on.
    const saysWhy = /SPFileCollection|document library/i.test(ghost.text || '');
    record('L2', 'POST /items on a document library',
           saysWhy ? 'REFUSED, AND SAYS WHY'
                   : isRefusal(ghost.status) ? 'REFUSED' : 'NOT ESTABLISHED',
           saysWhy
             ? `HTTP ${ghost.status}: ${ghost.text.slice(0, 300)}. This is the GOOD `
               + 'outcome, and the body names the reason rather than leaving the '
               + 'status to imply it: demo data cannot be generated for a library, '
               + 'and the build error that assumes so rests on this.'
             : isRefusal(ghost.status)
               ? `HTTP ${ghost.status}: ${ghost.text.slice(0, 300)}. Refused, but the `
                 + 'body does not say a library is why. Read it before treating this '
                 + 'as the library answer.'
               : `HTTP ${ghost.status}: ${ghost.text.slice(0, 300)}. That is not the `
                 + 'server refusing the content (a throttle or an outage lands here '
                 + 'too), so this run has not established what a library does with '
                 + 'a fileless POST. Re-run.');
    for (const id of ['L3', 'L4']) {
      record(id, RESULTS.find((r) => r.id === id).question, 'NOT APPLICABLE',
             'no item was created at L2');
    }
  } else {
    const id = ghost.body && ghost.body.Id;
    record('L2', 'POST /items on a document library', 'ACCEPTED',
           `HTTP ${ghost.status}, item ${id}. SharePoint did NOT refuse a fileless `
           + 'item, so the question becomes what it created (L3) and whether anyone '
           + 'can see it (L4)');
    const back = await spGet(
      `${listPath}/items(${id})?$select=Id,Title,FileSystemObjectType,FileRef,FileLeafRef`);
    const row = back.ok ? back.body : null;
    record('L3', 'What a fileless item POST actually created',
           row ? 'READ BACK' : 'NOT ESTABLISHED',
           row ? JSON.stringify(row) : `could not re-read item ${id} (HTTP ${back.status})`);

    // RenderListDataAsStream is what the modern UI itself calls, so this
    // answers "would staff see this" rather than "does REST list it".
    digest = await getDigest();
    const rendered = await spPost(
      `${listPath}/RenderListDataAsStream`,
      { parameters: { RenderOptions: 2 } }, digest);
    let visible = null;
    if (rendered.ok && rendered.body && Array.isArray(rendered.body.Row)) {
      visible = rendered.body.Row.some((r) => String(r.ID) === String(id));
    }
    record('L4', 'Whether that item is visible to a user',
           visible === null ? 'NOT ESTABLISHED' : visible ? 'VISIBLE' : 'HIDDEN FROM THE VIEW',
           visible === null
             ? `RenderListDataAsStream did not answer: HTTP ${rendered.status} `
               + rendered.text.slice(0, 200)
             : visible
               ? 'the fileless item appears in the default rendering, so staff would see '
                 + 'a document-shaped row with no document'
               : 'the item exists over REST but does not appear in the default '
                 + 'rendering, which is untidy rather than user-facing');
  }

  // ---- L5: does a real file populate Title? ---------------------------
  // The harness's spPost JSON-encodes its payload, which is wrong for a
  // file body, so this one call is a direct fetch.
  digest = await getDigest();
  const folder = `${LIB}`;
  let uploaded = null;
  try {
    const res = await fetch(
      `${WEB}/_api/web/GetFolderByServerRelativeUrl('${folder}')`
      + "/Files/add(url='dbmlsp-probe.txt',overwrite=true)",
      {
        method: 'POST',
        headers: {
          Accept: 'application/json;odata=nometadata',
          'X-RequestDigest': digest,
        },
        body: 'dbml-sharepoint probe file. Safe to delete.',
      });
    uploaded = { ok: res.ok, status: res.status, text: await res.text() };
  } catch (err) {
    uploaded = { ok: false, status: 0, text: String(err) };
  }
  if (!uploaded.ok) {
    record('L5', 'Is Title populated after a real file upload', 'NOT ESTABLISHED',
           `upload failed: HTTP ${uploaded.status} ${uploaded.text.slice(0, 260)}`);
  } else {
    const items = await spGet(
      `${listPath}/items?$select=Id,Title,FileLeafRef,FileSystemObjectType`
      + `&$filter=FileLeafRef eq 'dbmlsp-probe.txt'`);
    const row = items.ok && items.body && items.body.value ? items.body.value[0] : null;
    const title = row ? row.Title : undefined;
    const empty = title === null || title === '' || title === undefined;
    record('L5', 'Is Title populated after a real file upload',
           row ? (empty ? 'TITLE IS EMPTY, assumption HOLDS'
                        : 'TITLE IS POPULATED, assumption WRONG')
               : 'NOT ESTABLISHED',
           row
             ? `${JSON.stringify(row)}. `
               + (empty
                 ? 'the standard header line [$Title] would read empty on every '
                   + 'document, so a library header must reference something else '
                   + '(see L7).'
                 : 'SharePoint DID populate Title from the file, so the standard '
                   + 'header may be shippable on a library unchanged.')
             : `could not find the uploaded file's item (HTTP ${items.status})`);
  }

  // ---- L6: FileLeafRef in a view --------------------------------------
  // The tool refuses this at build time. The question is whether that is a
  // tool decision that could be revisited, or a platform refusal.
  const VIEW = 'dbmlsp probe docview';
  digest = await getDigest();
  const madeView = await spPost(`${listPath}/views`, {
    Title: VIEW, ViewQuery: '', RowLimit: 30,
  }, digest);
  if (!madeView.ok) {
    record('L6', 'Can a library view carry FileLeafRef through REST', 'NOT ESTABLISHED',
           `could not create a view: HTTP ${madeView.status} ${madeView.text.slice(0, 220)}`);
  } else {
    digest = await getDigest();
    const added = await spPost(
      `${listPath}/views/getbytitle('${VIEW}')/viewfields/addviewfield('FileLeafRef')`,
      {}, digest);
    const back = await spGet(
      `${listPath}/views/getbytitle('${VIEW}')/viewfields`);
    // A failed READ-BACK is not a refusal. Conflating them would let a
    // throttled GET print "REFUSED" (a claim about the platform) on no
    // observation at all, and report() would count the question answered.
    if (readFailed(back)) {
      record('L6', 'Can a library view carry FileLeafRef through REST', 'NOT ESTABLISHED',
             `addviewfield('FileLeafRef') returned HTTP ${added.status}, but the `
             + `view-fields read-back failed (HTTP ${back.status}), so this run has `
             + 'no evidence either way.');
    } else {
      const fields = back.body.Items || back.body.value || [];
      const present = JSON.stringify(fields).includes('FileLeafRef');
      record('L6', 'Can a library view carry FileLeafRef through REST',
             !added.ok ? 'REFUSED'
                       : present ? 'PLATFORM ALLOWS IT' : 'ACCEPTED THEN DISCARDED',
             !added.ok
               ? `addviewfield was refused with HTTP ${added.status}: `
                 + added.text.slice(0, 220)
               : present
                 ? `FileLeafRef is in the view's fields: ${JSON.stringify(fields)}. The `
                   + "tool's build error is a TOOL limit, and lifting it is a design "
                   + 'decision rather than a platform impossibility.'
                 : 'addviewfield returned HTTP ' + added.status + ' but the view reads '
                   + `back without FileLeafRef: ${JSON.stringify(fields)}`);
    }
  }

  // ---- L7: a header that names the file -------------------------------
  // Storage only. Whether the header RENDERS is the checklist's job below.
  const header = {
    elmType: 'div',
    attributes: { class: 'ms-borderColor-neutralTertiary' },
    children: [{
      elmType: 'div',
      attributes: { class: 'ms-fontColor-neutralSecondary ms-fontWeight-bold ms-fontSize-16' },
      txtContent: "=if([$FileLeafRef] == '', 'New document', 'Document: ' + [$FileLeafRef])",
    }],
  };
  // The collection is unordered and a library carries a Folder content type
  // as well as a Document one, so value[0] can be the folder, which would
  // store the formatter somewhere no document ever reads it and report STORED
  // regardless. Same rule the deploy uses (src/dbml_sharepoint/templates/
  // deploy/_forms.js.j2): the first id under 0x01 that is not under 0x0120.
  // The accessor stays `Id.StringValue`, not `StringId`: this harness runs
  // odata=nometadata, and StringValue is the spelling the 2026-07-29 run
  // proved works under it.
  const cts = await spGet(`${listPath}/contenttypes?$select=Id,Name&$top=20`);
  const ctIdOf = (ct) => String((ct && ct.Id && ct.Id.StringValue) || (ct && ct.Id) || '');
  const ct = ((cts.ok && cts.body && cts.body.value) || [])
    .find((c) => ctIdOf(c).startsWith('0x01') && !ctIdOf(c).startsWith('0x0120'));
  const ctId = ct ? ctIdOf(ct) : null;
  if (!ctId) {
    record('L7', 'A library content type accepts a header referencing [$FileLeafRef]',
           'NOT ESTABLISHED',
           `could not find a non-folder content type on the library (HTTP ${cts.status})`);
  } else {
    digest = await getDigest();
    const set = await spPost(`${listPath}/contenttypes('${ctId}')`, {
      ClientFormCustomFormatter: JSON.stringify({ header }),
    }, digest, { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
    if (!set.ok) {
      record('L7', 'A library content type accepts a header referencing [$FileLeafRef]',
             'REFUSED', `HTTP ${set.status}: ${set.text.slice(0, 260)}`);
    } else {
      const back = await spGet(
        `${listPath}/contenttypes('${ctId}')?$select=ClientFormCustomFormatter`);
      // Same rule as L6: a read-back that failed is not a discard.
      if (readFailed(back)) {
        record('L7', 'A library content type accepts a header referencing [$FileLeafRef]',
               'NOT ESTABLISHED',
               `the MERGE returned HTTP ${set.status}, but the read-back failed (HTTP `
               + `${back.status}), so whether it was stored is unobserved.`);
      } else {
        const stored = back.body.ClientFormCustomFormatter;
        record('L7', 'A library content type accepts a header referencing [$FileLeafRef]',
               stored && String(stored).includes('FileLeafRef')
                 ? 'STORED' : 'ACCEPTED THEN DISCARDED',
               `on content type '${ct.Name}' (${ctId}), reads back `
               + `${JSON.stringify(String(stored).slice(0, 220))}. STORAGE ONLY. `
               + 'See the checklist below.');
      }
    }
  }

  report();
  console.log('\n============ EYES-ON CHECKLIST ============');
  console.log(`  1. Open the library '${LIB}' and look at the default view.`);
  console.log('     Is there a row with no document behind it? (L4 predicts this;');
  console.log('     confirm what a user would actually see.)');
  console.log('     answer: ______________________________________');
  console.log('  2. Click the uploaded dbmlsp-probe.txt, open its details/edit');
  console.log('     panel, and say what the HEADER shows at the top.');
  console.log("     'New document', the file name, or nothing at all?");
  console.log('     answer: ______________________________________');
  console.log('\nThose two answers decided it: libraries are refused outright');
  console.log('until issue #14, and this is what a library header would have');
  console.log('to say if the kind ever comes back.');
  console.log('==========================================');
  log('INFO', `Done. Delete '${LIB}' when you have copied the results.`);
})();
