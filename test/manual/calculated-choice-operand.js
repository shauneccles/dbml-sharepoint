/**
 * dbml-sharepoint PROBE: CALCULATED COLUMN OVER CHOICE OPERANDS
 *
 * QUESTION: does SharePoint accept a calculated column whose operands are
 * two single-select Choice columns?
 *
 * WHY: templates/tiered-huddle declares
 *
 *     Route = '=[RaisedAtTier]&" -> "&[TargetTier]'
 *
 * where both operands are Choice. No first-party Microsoft source states
 * whether that is supported, and this project has been wrong about
 * unexercised SharePoint behaviour more than once. If it is refused,
 * deploy.js fails PART-WAY THROUGH provisioning, the same failure shape
 * website/docs/reference/dbml.md warns about for person and lookup
 * operands.
 *
 * WHAT IT ASKS
 *   C1  does the calculated column CREATE over two Choice operands?
 *   C2  does it RENDER a value on a saved item? (creating is not working)
 *   C3  what does SharePoint store in Formula, versus what was sent?
 *   C4  does a Choice value containing & " < survive concatenation?
 *   C5  what happens when one operand is blank?
 *
 *   D1  the same, but referencing operands by DISPLAY name WITH SPACES.
 *   D2  what is stored for that form? A bracketed name containing spaces
 *   D3  cannot have its brackets stripped without becoming ambiguous, so
 *       it may store and compare differently to C1/C3. This is the shape
 *       the tool actually emits: the build rewrites [RaisedAtTier] to
 *       [Raised At Tier] before deploying, so C1 alone tests a formula
 *       deploy.js never sends.
 *
 *   NUM1  ResultType Number over a Choice operand: accepted?
 *   NUM2  ...and does it compute the branch the Choice selects?
 *   DAT1  ResultType DateTime over Choice + Date operands: accepted?
 *   DAT2  ...and does the date offset compute?
 *       Both are declarable today (calculated_number, calculated_date) and
 *       are where this template goes next: a priority that sets a response
 *       time, or an escalation score derived from a choice.
 *
 *   R1  the retirement fold appends " (retired)" to a column's DISPLAY
 *   R2  title, and calculated formulas resolve operands BY display title.
 *       Does an existing formula survive that rename, and can a new one
 *       reference a title containing parentheses at all?
 *
 *   L1  is a Lookup operand refused in a calculated formula, as N1's
 *       Person operand is? The refusal message names no type list, so
 *       Lookup has to be asked rather than assumed.
 *   P1  a Person column in a COLUMN VALIDATION formula
 *   L2  a Lookup column in a COLUMN VALIDATION formula
 *   P2  a Person column in a CONDITIONAL VISIBILITY formula
 *   L3  a Lookup column in a CONDITIONAL VISIBILITY formula
 *       analysis/conditions.py forbids person in validation but permits
 *       lookup there, and permits both in conditional visibility, none of
 *       it evidenced. A rule that forbids what SharePoint allows is merely
 *       restrictive; one that PERMITS what SharePoint refuses fails at
 *       deploy time, part-way through provisioning.
 *
 *   N1  NEGATIVE CONTROL: is a Person operand refused?
 *
 * SECOND LIST: the lookup questions need something to point at, so this
 * also creates "<list> Target" with one row. CLEANUP removes both.
 *
 * FOR A CLEAN RUN set CLEANUP = true: a column that already exists is
 * reported as "already present", which is weaker evidence than actually
 * creating it.
 *
 * READ N1 FIRST. It is the only row that establishes this probe can tell
 * acceptance from refusal. If a Person operand is ACCEPTED, the probe is
 * not sensitive to failure and every other result is worthless.
 *
 * HOW TO RUN
 *   1. Open a site you own, at /_layouts/15/settings.aspx.
 *   2. F12 -> Console -> paste -> Enter. It prints its plan and stops.
 *   3. Edit CONFIRMED and ALLOW_WRITES to true, paste again.
 *   4. Copy the RESULTS block back verbatim.
 *
 * WHEN FINISHED: delete both lists it created. Everything lives in them.
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

  const LIST = 'dbmlsp Probe CalcChoice';
  const TIERS = ['Tier 1', 'Tier 2', 'Tier 3'];
  const PRIORITIES = ['High', 'Medium', 'Low'];
  // Deliberately awkward: & is SharePoint's concatenation operator, "
  // delimits its string literals, and < is significant in SchemaXml. If
  // concatenation mangles a Choice value, it mangles it here.
  const NASTY = 'A & B "quoted" <tag>';

  const xmlAttr = (s) => String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  if (!CONFIRMED) {
    log('INFO', `Would create list '${LIST}' on ${WEB}, add Choice, Date and`);
    log('INFO', 'Person and Lookup columns (the lookup needs a second list,');
    log('INFO', `'${LIST} Target'), then attempt calculated columns over them`);
    log('INFO', '(text, number and date results) and read back what was stored.');
    if (CLEANUP) {
      log('INFO', `CLEANUP is ON: '${LIST}' would be RECYCLED first, with its items.`);
    } else {
      log('INFO', `CLEANUP is off: an existing '${LIST}' would be topped up, and`);
      log('INFO', 'columns already present report "already present" rather than');
      log('INFO', 'being created. Set CLEANUP = true for a clean run.');
    }
    log('INFO', 'Nothing has been written. Set CONFIRMED and ALLOW_WRITES to true.');
    return;
  }
  if (!ALLOW_WRITES) {
    log('INFO', 'CONFIRMED, but ALLOW_WRITES is false and this probe must write.');
    log('INFO', 'Set ALLOW_WRITES = true to proceed. Stopping.');
    return;
  }

  // Declared before anything can fail, so an abort reports the questions
  // it never reached instead of only the ones it managed to ask.
  expect('C1', 'Calculated column over two Choice operands is accepted');
  expect('C2', 'Renders a value for two ordinary Choice values');
  expect('C3', 'Formula as SharePoint stored it');
  expect('C4', 'Renders a Choice value containing & " <');
  expect('C5', 'Behaviour when one operand is blank');
  expect('D1', 'Accepts operands referenced by DISPLAY name containing spaces');
  expect('D2', 'Spaced display-name formula as SharePoint stored it');
  expect('D3', 'Renders a value through spaced display-name operands');
  expect('NUM1', 'ResultType Number over a Choice operand is accepted');
  expect('NUM2', 'Number result computes the branch the Choice selects');
  expect('DAT1', 'ResultType DateTime over Choice + Date operands is accepted');
  expect('DAT2', 'Date result computes the offset the Choice selects');
  expect('R1', 'An existing calculated column survives its operand being re-titled "(retired)"');
  expect('R2', 'A NEW calculated column can reference a display name containing "(retired)"');
  expect('L1', 'A Lookup operand in a CALCULATED formula');
  expect('P1', 'A Person column in a COLUMN VALIDATION formula');
  expect('L2', 'A Lookup column in a COLUMN VALIDATION formula');
  expect('P2', 'A Person column in a CONDITIONAL VISIBILITY formula');
  expect('L3', 'A Lookup column in a CONDITIONAL VISIBILITY formula');
  expect('N1', 'NEGATIVE CONTROL: a Person operand is refused');

  // Removes a previous run's lists so every question below is answered by
  // actually creating something. No-op unless CLEANUP is on.
  // Main list FIRST: it holds the lookup column, and SharePoint refuses to
  // delete a list that a lookup still points at.
  await resetList(LIST);
  await resetList(`${LIST} Target`);

  let digest = await getDigest();

  // ---- Bootstrap. Every step checks first, so re-running tops up ------
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
  } else {
    log('INFO', `List '${LIST}' already exists, topping up.`);
  }

  const fieldsPath = `web/lists/getbytitle('${LIST}')/fields`;
  const itemsPath = `web/lists/getbytitle('${LIST}')/items`;

  const addField = async (schemaXml) => {
    digest = await getDigest();
    // No __metadata: the harness sends odata=nometadata, and that format
    // REJECTS the type hint rather than ignoring it:
    //   "The property '__metadata' does not exist on type
    //    'SP.XmlSchemaFieldCreationInformation'"
    // The odata=verbose examples in Microsoft's docs all carry it, which is
    // where this gets copied from.
    return spPost(`${fieldsPath}/createfieldasxml`, {
      parameters: {
        SchemaXml: schemaXml,
        Options: 8,  // AddFieldInternalNameHint
      },
    }, digest);
  };

  const fieldExists = async (name) =>
    (await spGet(`${fieldsPath}/getbyinternalnameortitle('${name}')`)).ok;

  // DisplayName and Name differ deliberately on the spaced columns. The
  // build rewrites [InternalName] to [Display Name] before deploying, and
  // auto_display_name turns RaisedAtTier into "Raised At Tier", so the
  // formula SharePoint actually receives references a name with SPACES.
  // Columns whose display and internal names match cannot exercise that.
  const choiceXml = (internal, display, choices) =>
    `<Field Type="Choice" DisplayName="${xmlAttr(display)}" Name="${internal}" Format="Dropdown">` +
    `<CHOICES>${choices.map((c) => `<CHOICE>${xmlAttr(c)}</CHOICE>`).join('')}` +
    `</CHOICES></Field>`;

  const bootstrap = [
    ['RaisedAtTier', 'RaisedAtTier', choiceXml('RaisedAtTier', 'RaisedAtTier', [...TIERS, NASTY])],
    ['TargetTier', 'TargetTier', choiceXml('TargetTier', 'TargetTier', [...TIERS, NASTY])],
    // Spaced display names, the shape the tool actually emits.
    ['SpacedFrom', 'Spaced From Tier', choiceXml('SpacedFrom', 'Spaced From Tier', TIERS)],
    ['SpacedTo', 'Spaced To Tier', choiceXml('SpacedTo', 'Spaced To Tier', TIERS)],
    // Drives a numeric score and a date offset, the two future shapes.
    ['ProbePriority', 'Probe Priority', choiceXml('ProbePriority', 'Probe Priority', PRIORITIES)],
    ['ProbeRaised', 'Probe Raised',
     '<Field Type="DateTime" DisplayName="Probe Raised" Name="ProbeRaised" Format="DateOnly"/>'],
    ['ProbeOwner', 'ProbeOwner',
     '<Field Type="User" DisplayName="ProbeOwner" Name="ProbeOwner" UserSelectionMode="PeopleOnly"/>'],
    // Retitled later to "Retire Me (retired)" to mimic the retirement fold.
    ['RetireMe', 'Retire Me', choiceXml('RetireMe', 'Retire Me', TIERS)],
  ];
  for (const [internal, , xml] of bootstrap) {
    if (!(await fieldExists(internal))) {
      const made = await addField(xml);
      if (!made.ok) {
        record('BOOT', `Create column ${internal}`, 'FAIL',
               `HTTP ${made.status}: ${made.text.slice(0, 300)}`);
        return report();
      }
    }
  }

  // ---- C1: the direct question ---------------------------------------
  // Display names equal internal names on this probe list, so the formula
  // needs no rewrite. jsgen does that display-name translation in a real
  // build; here it would only add a second variable to a single question.
  const ROUTE_FORMULA = '=[RaisedAtTier]&" -> "&[TargetTier]';
  const calcXml = (name, formula, refs, resultType = 'Text', extra = '') =>
    `<Field Type="Calculated" DisplayName="${name}" Name="${name}" ` +
    `ResultType="${resultType}"${extra}>` +
    `<Formula>${xmlAttr(formula)}</Formula>` +
    `<FieldRefs>${refs.map((r) => `<FieldRef Name="${r}"/>`).join('')}</FieldRefs>` +
    `</Field>`;

  if (!(await fieldExists('ProbeRoute'))) {
    const made = await addField(
      calcXml('ProbeRoute', ROUTE_FORMULA, ['RaisedAtTier', 'TargetTier']));
    record('C1', 'Calculated column over two Choice operands is accepted',
           made.ok ? 'PASS' : 'FAIL',
           made.ok
             ? `HTTP ${made.status} on createfieldasxml`
             : `HTTP ${made.status}: ${made.text.slice(0, 400)}`);
  } else {
    record('C1', 'Calculated column over two Choice operands is accepted', 'PASS',
           'column already present from an earlier run of this probe');
  }

  // ---- C3: what did SharePoint actually store? ------------------------
  const route = await spGet(`${fieldsPath}/getbyinternalnameortitle('ProbeRoute')`);
  if (route.ok) {
    record('C3', 'Formula as SharePoint stored it', 'INFO',
           `sent ${JSON.stringify(ROUTE_FORMULA)} / ` +
           `stored ${JSON.stringify(route.body.Formula)}`);
  } else {
    record('C3', 'Formula as SharePoint stored it', 'NOT ESTABLISHED',
           'the calculated column does not exist to read back');
  }

  // ---- C2 / C4 / C5: does it RENDER? ----------------------------------
  const cases = [
    ['C2', 'Renders a value for two ordinary Choice values',
     { Title: 'c2', RaisedAtTier: 'Tier 1', TargetTier: 'Tier 3' }, 'Tier 1 -> Tier 3'],
    ['C4', 'Renders a Choice value containing & " <',
     { Title: 'c4', RaisedAtTier: NASTY, TargetTier: 'Tier 2' }, `${NASTY} -> Tier 2`],
    ['C5', 'Behaviour when one operand is blank',
     { Title: 'c5', RaisedAtTier: 'Tier 1' }, null],
  ];

  if (!route.ok) {
    for (const [id, question] of cases) {
      record(id, question, 'NOT ESTABLISHED', 'no calculated column to render');
    }
  } else {
    for (const [id, question, payload, expected] of cases) {
      digest = await getDigest();
      const made = await spPost(itemsPath, payload, digest);
      if (!made.ok) {
        record(id, question, 'FAIL',
               `item create HTTP ${made.status}: ${made.text.slice(0, 300)}`);
        continue;
      }
      const back = await spGet(`${itemsPath}(${made.body.Id})?$select=ProbeRoute`);
      const actual = back.ok ? back.body.ProbeRoute : null;
      if (expected === null) {
        // No expected value: a blank operand's behaviour is the finding.
        record(id, question, 'INFO', `stored value: ${JSON.stringify(actual)}`);
      } else {
        record(id, question, actual === expected ? 'PASS' : 'FAIL',
               `expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
      }
    }
  }

  // ---- D: the shape the tool ACTUALLY emits ---------------------------
  // deploy.js carries =[Raised At Tier]&" -> "&[Target Tier], not the
  // internal-name form C1 tested. A bracketed name containing spaces
  // cannot have its brackets stripped without becoming ambiguous, so both
  // acceptance AND the stored form may differ from C1/C3.
  const SPACED_FORMULA = '=[Spaced From Tier]&" -> "&[Spaced To Tier]';
  if (!(await fieldExists('SpacedRoute'))) {
    const made = await addField(
      calcXml('SpacedRoute', SPACED_FORMULA, ['SpacedFrom', 'SpacedTo']));
    record('D1', 'Accepts operands referenced by DISPLAY name containing spaces',
           made.ok ? 'PASS' : 'FAIL',
           made.ok ? `HTTP ${made.status}` : `HTTP ${made.status}: ${made.text.slice(0, 400)}`);
  } else {
    record('D1', 'Accepts operands referenced by DISPLAY name containing spaces',
           'PASS', 'already present from an earlier run');
  }
  const spaced = await spGet(`${fieldsPath}/getbyinternalnameortitle('SpacedRoute')`);
  if (spaced.ok) {
    record('D2', 'Spaced display-name formula as SharePoint stored it', 'INFO',
           `sent ${JSON.stringify(SPACED_FORMULA)} / stored ${JSON.stringify(spaced.body.Formula)}`);
  }

  // ---- NUM / DAT: a Choice driving a score and a due date -------------
  // Both are declarable today (calculated_number, calculated_date) and are
  // the natural next step for this template: a priority that sets a
  // response time, or an escalation score.
  const NUM_FORMULA = '=IF([Probe Priority]="High",3,IF([Probe Priority]="Medium",2,1))';
  if (!(await fieldExists('ProbeScore'))) {
    const made = await addField(
      calcXml('ProbeScore', NUM_FORMULA, ['ProbePriority'], 'Number'));
    record('NUM1', 'ResultType Number over a Choice operand is accepted',
           made.ok ? 'PASS' : 'FAIL',
           made.ok ? `HTTP ${made.status}` : `HTTP ${made.status}: ${made.text.slice(0, 400)}`);
  } else {
    record('NUM1', 'ResultType Number over a Choice operand is accepted', 'PASS',
           'already present from an earlier run');
  }

  const DATE_FORMULA = '=[Probe Raised]+IF([Probe Priority]="High",1,7)';
  if (!(await fieldExists('ProbeDue'))) {
    const made = await addField(
      calcXml('ProbeDue', DATE_FORMULA, ['ProbeRaised', 'ProbePriority'], 'DateTime',
              ' Format="DateOnly"'));
    record('DAT1', 'ResultType DateTime over Choice + Date operands is accepted',
           made.ok ? 'PASS' : 'FAIL',
           made.ok ? `HTTP ${made.status}` : `HTTP ${made.status}: ${made.text.slice(0, 400)}`);
  } else {
    record('DAT1', 'ResultType DateTime over Choice + Date operands is accepted', 'PASS',
           'already present from an earlier run');
  }

  // One item exercises D3, NUM2 and DAT2 together.
  digest = await getDigest();
  const combo = await spPost(itemsPath, {
    Title: 'combo',
    SpacedFrom: 'Tier 1', SpacedTo: 'Tier 3',
    ProbePriority: 'High', ProbeRaised: '2026-03-02T00:00:00Z',
  }, digest);
  if (!combo.ok) {
    for (const id of ['D3', 'NUM2', 'DAT2']) {
      record(id, RESULTS.find((r) => r.id === id).question, 'FAIL',
             `item create HTTP ${combo.status}: ${combo.text.slice(0, 300)}`);
    }
  } else {
    const read = await spGet(
      `${itemsPath}(${combo.body.Id})?$select=SpacedRoute,ProbeScore,ProbeDue`);
    const got = read.ok ? read.body : {};
    record('D3', 'Renders a value through spaced display-name operands',
           got.SpacedRoute === 'Tier 1 -> Tier 3' ? 'PASS' : 'FAIL',
           `expected "Tier 1 -> Tier 3", got ${JSON.stringify(got.SpacedRoute)}`);
    record('NUM2', 'Number result computes the branch the Choice selects',
           Number(got.ProbeScore) === 3 ? 'PASS' : 'FAIL',
           `Priority "High" should score 3, got ${JSON.stringify(got.ProbeScore)}`);
    // Compared on the date part only: the stored value carries a timezone
    // and an exact-string match would fail for a reason that is not the
    // question being asked.
    const due = String(got.ProbeDue || '').slice(0, 10);
    record('DAT2', 'Date result computes the offset the Choice selects',
           due === '2026-03-03' ? 'PASS' : 'FAIL',
           `raised 2026-03-02 + High(1 day) should be 2026-03-03, got ` +
           `${JSON.stringify(got.ProbeDue)} (date part ${JSON.stringify(due)})`);
  }

  // ---- R: retirement re-titles an operand -----------------------------
  // The retirement fold appends " (retired)" to a column's DISPLAY title,
  // and calculated formulas resolve operands BY display title. Two
  // separate questions: does the existing formula survive the rename, and
  // can a new formula reference a title containing parentheses at all?
  if (!(await fieldExists('RetireRoute'))) {
    await addField(calcXml('RetireRoute', '=[Retire Me]&" fixed"', ['RetireMe']));
  }
  digest = await getDigest();
  const retitle = await spPost(
    `${fieldsPath}/getbyinternalnameortitle('RetireMe')`,
    { Title: 'Retire Me (retired)' },
    digest,
    { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
  );
  if (!retitle.ok) {
    record('R1', 'An existing calculated column survives its operand being re-titled "(retired)"',
           'NOT ESTABLISHED', `could not re-title: HTTP ${retitle.status}`);
  } else {
    digest = await getDigest();
    const row = await spPost(itemsPath, { Title: 'retire', RetireMe: 'Tier 2' }, digest);
    const readBack = row.ok
      ? await spGet(`${itemsPath}(${row.body.Id})?$select=RetireRoute`)
      : { ok: false };
    const after = await spGet(`${fieldsPath}/getbyinternalnameortitle('RetireRoute')`);
    record('R1', 'An existing calculated column survives its operand being re-titled "(retired)"',
           readBack.ok && readBack.body.RetireRoute === 'Tier 2 fixed' ? 'PASS' : 'FAIL',
           `after rename the stored formula is ${JSON.stringify(after.ok ? after.body.Formula : null)}; ` +
           `computed value ${JSON.stringify(readBack.ok ? readBack.body.RetireRoute : null)}`);

    const newRef = await addField(
      calcXml('RetiredRef', '=[Retire Me (retired)]&" x"', ['RetireMe']));
    record('R2', 'A NEW calculated column can reference a display name containing "(retired)"',
           newRef.ok ? 'PASS' : 'FAIL',
           newRef.ok ? `HTTP ${newRef.status}` : `HTTP ${newRef.status}: ${newRef.text.slice(0, 300)}`);
  }

  // ---- P / L: person and lookup, in all three formula stores ----------
  // The tool's own rules here are uneven, and none of them were evidenced:
  //   analysis/conditions.py forbids `person` in a VALIDATION formula via
  //   _FORBIDDEN_OPERAND_TYPES, and forbids `lookup` there via a SEPARATE
  //   guard (_lookup_problem). A lookup is int-typed in DBML, so the type
  //   map alone cannot see it. Both are permitted in an EXPRESSION
  //   (conditional visibility) formula.
  // A rule that forbids something SharePoint allows is merely restrictive.
  // A rule that PERMITS something SharePoint refuses fails at deploy time,
  // part-way through provisioning, so L2, P2 and L3 are the ones that
  // matter most.
  const targetList = `${LIST} Target`;
  let lookupReady = false;
  const targetProbe = await spGet(`web/lists/getbytitle('${targetList}')`);
  if (!targetProbe.ok) {
    digest = await getDigest();
    await spPost('web/lists', {
      Title: targetList, BaseTemplate: 100,
      Description: 'dbml-sharepoint probe lookup target. Safe to delete.',
    }, digest);
  }
  const target = await spGet(`web/lists/getbytitle('${targetList}')`);
  if (target.ok) {
    digest = await getDigest();
    await spPost(`web/lists/getbytitle('${targetList}')/items`, { Title: 'row one' }, digest);
    if (!(await fieldExists('ProbeLookup'))) {
      await addField(
        `<Field Type="Lookup" DisplayName="ProbeLookup" Name="ProbeLookup" ` +
        `List="{${target.body.Id}}" ShowField="Title"/>`);
    }
    lookupReady = await fieldExists('ProbeLookup');
  }

  // L1: a Lookup operand in a calculated formula. N1 establishes that a
  // Person operand is refused, but the error names no type list, so Lookup
  // has to be asked separately rather than assumed to behave the same.
  if (!lookupReady) {
    record('L1', 'A Lookup operand in a CALCULATED formula', 'NOT ESTABLISHED',
           'the lookup column could not be created');
  } else if (!(await fieldExists('LookupCalc'))) {
    const made = await addField(calcXml('LookupCalc', '=[ProbeLookup]&" x"', ['ProbeLookup']));
    record('L1', 'A Lookup operand in a CALCULATED formula',
           made.ok ? 'ACCEPTED' : 'REFUSED',
           made.ok
             ? `HTTP ${made.status}: SharePoint ALLOWS this; the README says it does not`
             : `HTTP ${made.status}: ${made.text.slice(0, 300)}`);
  } else {
    record('L1', 'A Lookup operand in a CALCULATED formula', 'ACCEPTED',
           'already present from an earlier run');
  }

  // Both validation stores live on the field and are set by MERGE.
  // ValidationFormula takes DOUBLE-quoted literals and AND()/OR();
  // ClientValidationFormula takes SINGLE quotes and &&/||, established
  // earlier by form-visibility-evidence-probe.js.
  const setOnField = async (field, body) => {
    digest = await getDigest();
    return spPost(`${fieldsPath}/getbyinternalnameortitle('${field}')`, body, digest,
                  { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
  };
  const readBackOf = async (field, prop) => {
    const r = await spGet(`${fieldsPath}/getbyinternalnameortitle('${field}')?$select=${prop}`);
    return r.ok ? r.body[prop] : null;
  };

  const stores = [
    ['P1', 'A Person column in a COLUMN VALIDATION formula', 'ProbeOwner',
     { ValidationFormula: '=NOT(ISBLANK([ProbeOwner]))' }, 'ValidationFormula',
     'the tool FORBIDS this today'],
    ['L2', 'A Lookup column in a COLUMN VALIDATION formula', 'ProbeLookup',
     { ValidationFormula: '=NOT(ISBLANK([ProbeLookup]))' }, 'ValidationFormula',
     'the tool FORBIDS this today, via _lookup_problem rather than the type map'],
    ['P2', 'A Person column in a CONDITIONAL VISIBILITY formula', 'ProbeOwner',
     { ClientValidationFormula: "=if([$ProbeOwner] != '', 'true', 'false')" },
     'ClientValidationFormula', 'the tool PERMITS this today'],
    ['L3', 'A Lookup column in a CONDITIONAL VISIBILITY formula', 'ProbeLookup',
     { ClientValidationFormula: "=if([$ProbeLookup] != '', 'true', 'false')" },
     'ClientValidationFormula', 'the tool PERMITS this today'],
  ];
  for (const [id, question, field, body, prop, toolStance] of stores) {
    if (field === 'ProbeLookup' && !lookupReady) {
      record(id, question, 'NOT ESTABLISHED', 'the lookup column could not be created');
      continue;
    }
    const set = await setOnField(field, body);
    // A store that accepts the write but discards it is the outcome that
    // would fool a deploy: read it back rather than trusting HTTP 200.
    const stored = set.ok ? await readBackOf(field, prop) : null;
    if (!set.ok) {
      record(id, question, 'REFUSED',
             `${toolStance}; HTTP ${set.status}: ${set.text.slice(0, 300)}`);
    } else if (!stored) {
      record(id, question, 'ACCEPTED THEN DISCARDED',
             `${toolStance}; HTTP ${set.status} but ${prop} reads back ${JSON.stringify(stored)}`);
    } else {
      record(id, question, 'ACCEPTED',
             `${toolStance}; stored ${JSON.stringify(stored)}`);
    }
  }

  // ---- N1: NEGATIVE CONTROL -------------------------------------------
  // A Person operand in a calculated formula is documented as unsupported.
  // If this SUCCEEDS, this probe cannot distinguish acceptance from
  // refusal, and every PASS above is unproven rather than wrong.
  if (!(await fieldExists('ProbeNegative'))) {
    const negative = await addField(
      calcXml('ProbeNegative', '=[ProbeOwner]&" x"', ['ProbeOwner']));
    record('N1', 'NEGATIVE CONTROL: a Person operand is refused',
           negative.ok ? 'FAIL' : 'PASS',
           negative.ok
             ? 'Person operand was ACCEPTED. This probe cannot detect a refusal, '
               + 'so treat every other row as unproven'
             : `refused with HTTP ${negative.status}: ${negative.text.slice(0, 300)}`);
  } else {
    record('N1', 'NEGATIVE CONTROL: a Person operand is refused', 'NOT ESTABLISHED',
           'ProbeNegative already exists, so this run did not test the refusal. '
           + 'Delete the list and re-run for a clean control.');
  }

  report();
  log('INFO', `Done. Delete '${LIST}' and '${LIST} Target' when you have copied the results.`);
})();
