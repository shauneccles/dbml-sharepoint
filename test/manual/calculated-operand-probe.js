/**
 * dbml-sharepoint PROBE: CALCULATED COLUMN OPERAND MATRIX
 *
 * QUESTION: which SharePoint column types may a calculated field reference?
 *
 * WHY: Microsoft documents Lookup fields as unsupported and lists the
 * supported scalar operand types, while this project had live evidence that
 * a Person operand is refused with HTTP 500. Long text, rich text and
 * hyperlink were ambiguous (absent from Microsoft's supported list, which is
 * not the same as documented against), so they were kept OUT of the
 * validator's denylist until this probe ran. See STATUS.
 *
 * SOURCE
 *   https://support.microsoft.com/en-us/sharepoint/lists/data-and-lists/examples-of-common-formulas-in-lists
 *
 * WHAT IT ASKS
 *   LOOK  Lookup        PERS  Person       LONG  plain multi-line text
 *   RICH  rich text     LINK  Hyperlink    BOOL  Yes/No
 *   CHOI  Choice        DATE  Date only    TIME  Date and time
 *   NUMB  Number        TEXT  single text  CALC  another calculated field
 *
 * Each row records the createfieldasxml HTTP status and SharePoint error
 * body. Creation is the relevant gate: the deployment currently dies there,
 * part-way through provisioning.
 *
 * HOW TO RUN
 *   1. Open a disposable SharePoint Online site you own.
 *   2. F12 -> Console -> paste -> Enter. The committed defaults only print.
 *   3. Set CONFIRMED and ALLOW_WRITES to true. Set CLEANUP to true for a
 *      clean run and CLEANUP_AT_END to true to recycle both probe lists.
 *   4. Paste again and copy the complete RESULTS block back verbatim.
 *
 * STATUS: RUN 2026-07-30 against a live SharePoint Online site. All twelve
 * questions answered, none left open. Verbatim outcome:
 *
 *   LOOK  REFUSED    Lookup
 *   PERS  REFUSED    Person
 *   LONG  REFUSED    plain multi-line text (Note, RichText="FALSE")
 *   RICH  REFUSED    rich text (Note, RichText="TRUE")
 *   LINK  REFUSED    Hyperlink (URL)
 *   BOOL  ACCEPTED   Yes/No
 *   CHOI  ACCEPTED   Choice
 *   DATE  ACCEPTED   Date only
 *   TIME  ACCEPTED   Date and time
 *   NUMB  ACCEPTED   Number
 *   TEXT  ACCEPTED   single line of text
 *   CALC  ACCEPTED   another calculated column
 *
 * Every refusal was HTTP 500 with one identical body:
 *
 *   {"odata.error":{"code":"-2130575272, Microsoft.SharePoint.SPException",
 *    "message":{"lang":"en-US","value":"One or more column references are not
 *    allowed, because the columns are defined as a data type that is not
 *    supported in formulas."}}}
 *
 * All five refused types are now in _FORBIDDEN_CALCULATED_OPERANDS in
 * analysis/checks/_structure.py, so the build refuses them before a script is
 * emitted. Note what the run also showed: the three ambiguous types were
 * refused, so the cautious guess would have been RIGHT, and that is not a
 * reason to guess next time. The same caution kept Yes/No out of the
 * denylist, where a guess would have been WRONG.
 *
 * Incidental, and worth knowing before reading a future run: a GET on
 * fields/getbyinternalnameortitle() for a field that does not exist answers
 * HTTP 400, not 404. fieldExists() treats any non-2xx as absent, so this is
 * already handled, but a reader scanning the console for 404s will not find
 * them.
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

  const LIST = 'dbmlsp Probe CalcOperands';
  const TARGET = `${LIST} Target`;
  // Ships off so a pasted, unedited probe never removes anything. Turn it on
  // with the two write gates to recycle the two probe-owned lists after the
  // result table has printed.
  const CLEANUP_AT_END = false;

  const QUESTIONS = [
    ['LOOK', 'Lookup operand in a calculated formula'],
    ['PERS', 'Person operand in a calculated formula'],
    ['LONG', 'Plain multi-line-text operand in a calculated formula'],
    ['RICH', 'Rich-text operand in a calculated formula'],
    ['LINK', 'Hyperlink operand in a calculated formula'],
    ['BOOL', 'Yes/No operand in a calculated formula'],
    ['CHOI', 'Choice operand in a calculated formula'],
    ['DATE', 'Date-only operand in a calculated formula'],
    ['TIME', 'Date-and-time operand in a calculated formula'],
    ['NUMB', 'Number operand in a calculated formula'],
    ['TEXT', 'Single-line-text operand in a calculated formula'],
    ['CALC', 'Calculated-column operand in another calculated formula'],
  ];
  // Literal registrations are deliberate: test_probes statically proves that
  // every record('ID', ...) of a DECLARED QUESTION has a matching upfront
  // expect('ID', ...), so an aborted run cannot make unanswered questions
  // disappear. BOOT-prefixed ids are exempt there by design. They report a
  // bootstrap failure rather than answering a question, so there is no
  // question for them to hide.
  expect('LOOK', 'Lookup operand in a calculated formula');
  expect('PERS', 'Person operand in a calculated formula');
  expect('LONG', 'Plain multi-line-text operand in a calculated formula');
  expect('RICH', 'Rich-text operand in a calculated formula');
  expect('LINK', 'Hyperlink operand in a calculated formula');
  expect('BOOL', 'Yes/No operand in a calculated formula');
  expect('CHOI', 'Choice operand in a calculated formula');
  expect('DATE', 'Date-only operand in a calculated formula');
  expect('TIME', 'Date-and-time operand in a calculated formula');
  expect('NUMB', 'Number operand in a calculated formula');
  expect('TEXT', 'Single-line-text operand in a calculated formula');
  expect('CALC', 'Calculated-column operand in another calculated formula');

  if (!CONFIRMED) {
    log('INFO', `Would create '${TARGET}' and '${LIST}', add one source column`);
    log('INFO', 'of every supported tool type, then attempt one calculated');
    log('INFO', 'field over each source and print every HTTP response.');
    log('INFO', 'Nothing has been written. Set CONFIRMED and ALLOW_WRITES to true.');
    return;
  }
  if (!ALLOW_WRITES) {
    log('INFO', 'CONFIRMED, but ALLOW_WRITES is false. Stopping without writes.');
    return;
  }

  // Main first because it owns the lookup into TARGET.
  await resetList(LIST);
  await resetList(TARGET);

  let digest = await getDigest();
  // bootId is per LIST, not a shared 'BOOT'. record() overwrites by id, so one
  // id for both lists means whichever fails second erases the first, and the
  // surviving row names the wrong list in its own question text. Two lists
  // bootstrap here, so two ids.
  const ensureList = async (title, bootId) => {
    const existing = await spGet(`web/lists/getbytitle('${title}')?$select=Id`);
    if (existing.ok) return existing.body;
    digest = await getDigest();
    const created = await spPost('web/lists', {
      Title: title,
      BaseTemplate: 100,
      Description: 'dbml-sharepoint calculated-operand probe. Safe to recycle.',
    }, digest);
    if (!created.ok) {
      record(bootId, `Create probe list ${title}`, 'FAIL',
             `HTTP ${created.status}: ${created.text.slice(0, 400)}`);
      return null;
    }
    // SharePoint's response shape varies with OData mode and can be empty
    // after a successful create. Re-read the list so the Lookup schema below
    // always receives a measured Id rather than trusting the POST payload.
    const reread = await spGet(`web/lists/getbytitle('${title}')?$select=Id`);
    if (!reread.ok || !reread.body || !reread.body.Id) {
      record(bootId, `Read back probe list ${title}`, 'FAIL',
             `HTTP ${reread.status}: successful create returned no usable list Id`);
      return null;
    }
    return reread.body;
  };

  // Both are attempted even if the first fails, so one run reports the state
  // of both lists rather than only the one it reached.
  const target = await ensureList(TARGET, 'BOOTTARGET');
  const main = await ensureList(LIST, 'BOOTMAIN');
  if (!target || !main) {
    report();
    return;
  }

  const fieldsPath = `web/lists/getbytitle('${LIST}')/fields`;
  const xmlAttr = (value) => String(value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const fieldExists = async (name) =>
    (await spGet(`${fieldsPath}/getbyinternalnameortitle('${name}')?$select=Id`)).ok;
  const addField = async (schemaXml) => {
    digest = await getDigest();
    return spPost(`${fieldsPath}/createfieldasxml`, {
      parameters: { SchemaXml: schemaXml, Options: 8 },
    }, digest);
  };

  const choiceXml =
    '<Field Type="Choice" DisplayName="ProbeChoice" Name="ProbeChoice" Format="Dropdown">' +
    '<CHOICES><CHOICE>Alpha</CHOICE><CHOICE>Beta</CHOICE></CHOICES></Field>';
  const sources = [
    ['ProbeLookup',
     `<Field Type="Lookup" DisplayName="ProbeLookup" Name="ProbeLookup" ` +
     `List="{${target.Id}}" ShowField="Title"/>`],
    ['ProbePerson',
     '<Field Type="User" DisplayName="ProbePerson" Name="ProbePerson" ' +
     'UserSelectionMode="PeopleOnly"/>'],
    ['ProbeLong',
     '<Field Type="Note" DisplayName="ProbeLong" Name="ProbeLong" ' +
     'RichText="FALSE" NumLines="6"/>'],
    ['ProbeRich',
     '<Field Type="Note" DisplayName="ProbeRich" Name="ProbeRich" ' +
     'RichText="TRUE" RichTextMode="FullHtml" NumLines="6"/>'],
    ['ProbeLink',
     '<Field Type="URL" DisplayName="ProbeLink" Name="ProbeLink" Format="Hyperlink"/>'],
    ['ProbeBool',
     '<Field Type="Boolean" DisplayName="ProbeBool" Name="ProbeBool"/>'],
    ['ProbeChoice', choiceXml],
    ['ProbeDate',
     '<Field Type="DateTime" DisplayName="ProbeDate" Name="ProbeDate" Format="DateOnly"/>'],
    ['ProbeTime',
     '<Field Type="DateTime" DisplayName="ProbeTime" Name="ProbeTime" Format="DateTime"/>'],
    ['ProbeNumber',
     '<Field Type="Number" DisplayName="ProbeNumber" Name="ProbeNumber"/>'],
    ['ProbeText',
     '<Field Type="Text" DisplayName="ProbeText" Name="ProbeText" MaxLength="255"/>'],
  ];
  const sourceReady = new Set();
  for (const [name, schemaXml] of sources) {
    if (await fieldExists(name)) {
      sourceReady.add(name);
      continue;
    }
    const made = await addField(schemaXml);
    if (made.ok) {
      sourceReady.add(name);
    } else {
      log('FAIL', `Could not create source ${name}: HTTP ${made.status} ${made.text.slice(0, 300)}`);
    }
  }

  const calcXml = (name, formula, refs, resultType = 'Text') =>
    `<Field Type="Calculated" DisplayName="${name}" Name="${name}" ` +
    `ResultType="${resultType}">` +
    `<Formula>${xmlAttr(formula)}</Formula>` +
    `<FieldRefs>${refs.map((ref) => `<FieldRef Name="${ref}"/>`).join('')}</FieldRefs>` +
    '</Field>';

  const attempts = [
    ['LOOK', 'ProbeLookup', 'CalcLookup', '=[ProbeLookup]', 'Text'],
    ['PERS', 'ProbePerson', 'CalcPerson', '=[ProbePerson]', 'Text'],
    ['LONG', 'ProbeLong', 'CalcLong', '=[ProbeLong]', 'Text'],
    ['RICH', 'ProbeRich', 'CalcRich', '=[ProbeRich]', 'Text'],
    ['LINK', 'ProbeLink', 'CalcLink', '=[ProbeLink]', 'Text'],
    ['BOOL', 'ProbeBool', 'CalcBool', '=IF([ProbeBool],"yes","no")', 'Text'],
    ['CHOI', 'ProbeChoice', 'CalcChoice', '=[ProbeChoice]', 'Text'],
    ['DATE', 'ProbeDate', 'CalcDate', '=[ProbeDate]', 'DateTime'],
    ['TIME', 'ProbeTime', 'CalcTime', '=[ProbeTime]', 'DateTime'],
    ['NUMB', 'ProbeNumber', 'CalcNumber', '=[ProbeNumber]', 'Number'],
    ['TEXT', 'ProbeText', 'CalcText', '=[ProbeText]', 'Text'],
  ];

  const questionFor = (id) => QUESTIONS.find(([candidate]) => candidate === id)[1];
  const attempt = async (id, source, output, formula, resultType) => {
    const question = questionFor(id);
    if (!sourceReady.has(source)) {
      record(id, question, 'NOT ESTABLISHED', `${source} could not be created`);
      return false;
    }
    if (await fieldExists(output)) {
      record(id, question, 'NOT ESTABLISHED',
             `${output} already exists; use CLEANUP=true for a real creation attempt`);
      return true;
    }
    const made = await addField(calcXml(output, formula, [source], resultType));
    record(
      id,
      question,
      made.ok ? 'ACCEPTED' : (isRefusal(made.status) ? 'REFUSED' : 'NOT ESTABLISHED'),
      `HTTP ${made.status}${made.text ? `: ${made.text.slice(0, 500)}` : ''}`,
    );
    return made.ok;
  };

  for (const row of attempts) await attempt(...row);

  // CALC needs an accepted calculated source. Number is documented as a
  // supported operand and produces a simple, type-stable base field.
  const baseReady = await fieldExists('CalcNumber');
  if (!baseReady) {
    record('CALC', questionFor('CALC'), 'NOT ESTABLISHED',
           'CalcNumber was not created, so no calculated source exists');
  } else {
    sourceReady.add('CalcNumber');
    await attempt('CALC', 'CalcNumber', 'CalcCalculated', '=[CalcNumber]', 'Number');
  }

  report();

  if (CLEANUP_AT_END) {
    // Main first: SharePoint may refuse to recycle a list still targeted by
    // a Lookup column on another list.
    for (const title of [LIST, TARGET]) {
      digest = await getDigest();
      const recycled = await spPost(
        `web/lists/getbytitle('${title}')/recycle`, {}, digest,
      );
      log(
        recycled.ok ? 'OK' : 'FAIL',
        recycled.ok
          ? `Recycled '${title}'. It is recoverable from the site recycle bin.`
          : `Could not recycle '${title}': HTTP ${recycled.status} ${recycled.text.slice(0, 300)}`,
      );
    }
  } else {
    log('INFO', `Probe lists remain: '${LIST}' and '${TARGET}'.`);
    log('INFO', 'After copying results, set CLEANUP_AT_END=true and rerun, or recycle them manually.');
  }
})();
