/**
 * dbml-sharepoint VIEW AGGREGATIONS PROBE (creates and deletes its own list).
 *
 * Answers one question the documentation does not, before `totals:` is
 * built into the tool: HOW is a view's totals row actually written, and
 * does writing it actually make a total appear?
 *
 * Three templates publish a totals view they cannot build: vehicle-log's
 * "Monthly km by vehicle" (sum TripKm), service-requests' turnaround
 * report and complaints-feedback's monthly report. `Aggregations` is a
 * property of SP.View rather than part of ViewQuery, and which write path
 * takes it is not established. This repository does not ship a write
 * mechanism it has not read back from a live list, so the feature is
 * gated on this probe rather than on an assumption.
 *
 * WHAT IT WRITES: one list, named by PROBE_LIST below, created at start
 * and deleted at end. It never reads, writes or enumerates any other
 * list. If cleanup fails it says so loudly and prints the URL to delete
 * by hand.
 *
 * HOW TO RUN
 *   1. Paste it once: it prints the web and stops. Set CONFIRMED = true.
 *   2. Open that site's classic settings page (/_layouts/15/settings.aspx),
 *      signed in as a Site Owner. The site guard needs _spPageContextInfo.
 *   3. F12 -> Console -> type `allow pasting` if the browser objects ->
 *      paste this whole file -> Enter.
 *   4. Read the RESULTS table, then do THE MANUAL STEP below.
 *   5. Paste the [VERDICT] line back to whoever asked for this probe.
 *
 * THE MANUAL STEP (Q4 cannot be answered from script)
 *   A property that round-trips but renders no totals row is
 *   indistinguishable from a working one on every check a deploy can
 *   perform, which is exactly the failure class this repository keeps
 *   closing. So set CLEANUP_AT_END = false, run, then OPEN THE VIEW URL
 *   the probe prints and LOOK for the totals row under the Amount column.
 *   Re-run with CLEANUP_AT_END = true to remove the list.
 *
 * THIRD RUN 2026-08-14, revision 96d0a67a, on a THIRD site. Every result
 * below reproduced again: seeded=ok, mechanism=patch (HTTP 204), readback
 * exact, Sum of 42 and Average of 2 both rendered, so internal names bind
 * and two aggregations render together. Three independent sites now agree,
 * and nothing has ever contradicted the 2026-07-29 answers.
 *
 * That run also confirmed the self-seeding fix below works. The operator
 * settled Q5 and Q6 from the seeded data alone, with none of the hand-typing
 * the aa79f6c4 run needed.
 *
 * RE-RUN 2026-08-14, revision aa79f6c4, on a different site. Every result
 * below REPRODUCED: mechanism=patch (HTTP 204), readback exact, a Sum of 42
 * rendered, internal names bound, both columns rendered in order.
 *
 * It also surfaced something the 2026-07-29 record does not mention, and it
 * matters more than it looks:
 *
 *   AN AGGREGATION OVER AN EMPTY COLUMN RENDERS NOTHING AT ALL. No label,
 *   no zero, no blank total. The column footer is simply absent. The probe
 *   created `SecondAmount` AFTER seeding its rows, so both were empty, and
 *   the operator running it saw exactly what a FAILED BINDING looks like.
 *   Q5 could only be settled by hand-typing values into the list, at which
 *   point "Average 2" appeared under the display title and the binding was
 *   proven all along.
 *
 *   That is a diagnostic trap for anyone deploying a totals view onto a NEW
 *   list: it will look broken until the first row carries a value. It is
 *   also why this probe now seeds the second column itself. A probe that
 *   cannot answer its own question without manual data entry is one whose
 *   answer depends on the operator guessing what to try.
 *
 * ANSWERED, 2026-07-29, against a live SharePoint Online site:
 *
 *   seeded=ok mechanism=patch readback=ok rendered=yes
 *   internal_names=bind  two_columns=render  order=preserved
 *
 *   A REST MERGE of SP.View with Aggregations and AggregationsStatus is
 *   accepted (HTTP 204), reads back as
 *   `<FieldRef Name="Amount" Type="Sum" />` with AggregationsStatus `On`,
 *   and the view renders a totals row. SetViewXml was never attempted
 *   because the simpler mechanism won. The probe is kept for re-running
 *   against a tenant whose behaviour is in doubt.
 *
 *   Q5: AGGREGATIONS BINDS BY INTERNAL NAME. A FieldRef naming
 *       `SecondAmount` produced a figure under its display title
 *       "Second Amount Display". This is the opposite of the sibling
 *       ColumnWidth property, which binds by DISPLAY title and silently
 *       resets when given internal names, so the two are written
 *       differently on purpose, and neither can be inferred from the
 *       other.
 *   Q6: TWO AGGREGATIONS BOTH RENDER, and the readback preserves
 *       declaration order:
 *       `<FieldRef Name="Amount" Type="Sum" /><FieldRef Name="SecondAmount" Type="AVG" />`
 *       The deployer compares that string exactly, so order mattering is
 *       the safe outcome. Tokens also round-trip verbatim (`Sum` stays
 *       `Sum` and `AVG` stays `AVG`), so SharePoint normalises neither
 *       case nor spelling, and the only difference to absorb is the
 *       pre-`/>` space.
 *
 *   THE AGGREGATION TYPE IS AN ENUMERATION, NOT ENGLISH: AVG, COUNT, MAX,
 *   MIN, SUM, STDEV, VAR, per FieldRef element (Query), case-insensitive.
 *   A non-member such as "Average" is ACCEPTED, stored and read back
 *   unchanged, and then the view will not render at all ("Unknown render
 *   failure"). That is worse than a silent no-op, and it is observed
 *   behaviour on this tenant, not a caution.
 *
 *   To recover a view in that state, MERGE it with Aggregations: '' and
 *   AggregationsStatus: 'Off'.
 *
 *   SEEDING IS Q0 (posted, counted, and read back from ItemCount), and a
 *   failed seed downgrades the verdict line rather than sitting quietly
 *   beneath it. An empty list shows no totals row whether the feature works
 *   or not, so a `rendered=no` from an unseeded run means nothing.
 */
(async () => {
  // ---- Operator settings -------------------------------------------------
  const CONFIRMED = false;
  const PROBE_LIST = 'zzz dbmlsp view aggregations probe';
  const CLEANUP_AT_END = false;
  // What the totals row should show once one of the attempts lands.
  const AGG_FIELD = 'Amount';
  const AGG_TYPE = 'Sum';
  // ------------------------------------------------------------------------

  const log = (level, msg) => console.log(`[SP-PROBE] [${level}] ${msg}`);

  // Printed before any gate: a stale clipboard and a fix that did not
  // work produce identical transcripts otherwise.
  log('INFO', 'probe revision b1c969d3. Quote this when reporting results.');
  const results = [];
  const expect = (id, question) => {
    results.push({ id, question, observed: 'NOT ESTABLISHED', detail: 'the run did not reach this question' });
  };
  const record = (id, question, observed, detail) => {
    const row = results.find((r) => r.id === id);
    if (row) {
      Object.assign(row, { question, observed, detail: detail || '' });
    } else {
      results.push({ id, question, observed, detail: detail || '' });
    }
    log('INFO', `${id}: ${observed}${detail ? ` (${detail})` : ''}`);
  };
  expect('Q0', 'two rows actually seeded, so the manual check has something to total');
  expect('Q1', 'REST PATCH of SP.View Aggregations/AggregationsStatus is accepted');
  expect('Q2', 'GetViewXml/SetViewXml carries an <Aggregations> block');
  expect('Q3', 'the written property reads back unchanged');
  expect('Q4', 'a totals row actually RENDERS (manual: open the view URL and look)');
  expect('Q5', 'Aggregations binds by INTERNAL name, not display title (manual: look)');
  expect('Q6', 'two totalled columns both render, in declaration order (manual: look)');

  // === Preflight: confirm the site ===
  // SP REST '/_api/...' is routed by the path prefix BEFORE '_api'. A bare
  // '/_api/web/...' targets the tenant root web, NOT the sub-site you are
  // viewing. Prefix every call with the current web's server-relative URL.
  if (typeof _spPageContextInfo === 'undefined') {
    log('ERROR', '_spPageContextInfo is not available on this page; cannot resolve the web context. Open /_layouts/15/settings.aspx and retry.');
    return { aborted: 'no-sp-page-context' };
  }
  const WEB = (_spPageContextInfo.webServerRelativeUrl || '').replace(/\/$/, '');
  if (!CONFIRMED) {
    log('INFO', `This page is ${window.location.origin}${WEB || '/'}.`);
    log('INFO', 'If that is the site you want, set CONFIRMED = true and paste again.');
    return { aborted: 'unconfirmed' };
  }
  const apiUrl = (suffix) => `${WEB}/_api/${suffix}`;
  const odataName = (name) => encodeURIComponent(String(name).replace(/'/g, "''"));
  log('INFO', `Running as ${_spPageContextInfo.userLoginName || '(unknown)'} on web '${WEB || '(root)'}'.`);

  // === Transport ===
  const sleep = (ms) => new Promise((res) => setTimeout(res, ms));
  const spError = (text) => {
    try {
      return JSON.parse(text)?.error?.message?.value || String(text).slice(0, 300);
    } catch {
      return String(text).slice(0, 300);
    }
  };
  async function fetchWithRetry(url, opts, attempts = 5) {
    for (let i = 0; ; i++) {
      const r = await fetch(url, opts);
      if ((r.status === 429 || r.status === 503) && i < attempts) {
        const ra = Number(r.headers.get('Retry-After')) || Math.min(2 ** i, 30);
        log('INFO', `Throttled (HTTP ${r.status}); retry ${i + 1}/${attempts} in ${ra}s.`);
        await sleep(ra * 1000);
        continue;
      }
      return r;
    }
  }
  let cachedDigest = null;
  let digestExpiresAt = 0;
  async function getDigest() {
    if (cachedDigest && Date.now() < digestExpiresAt) return cachedDigest;
    const r = await fetchWithRetry(apiUrl('contextinfo'), {
      method: 'POST', headers: { 'Accept': 'application/json;odata=verbose' },
    });
    const info = (await r.json()).d.GetContextWebInformation;
    cachedDigest = info.FormDigestValue;
    digestExpiresAt = Date.now() + Math.max((Number(info.FormDigestTimeoutSeconds) || 1800) - 60, 60) * 1000;
    return cachedDigest;
  }
  const spHeaders = (digest, extra = {}) => ({
    'Accept': 'application/json;odata=verbose',
    'Content-Type': 'application/json;odata=verbose',
    'X-RequestDigest': digest,
    ...extra,
  });
  async function post(suffix, body, extraHeaders) {
    const digest = await getDigest();
    const r = await fetchWithRetry(apiUrl(suffix), {
      method: 'POST',
      headers: spHeaders(digest, extraHeaders || {}),
      body: body === undefined ? undefined : JSON.stringify(body),
    });
    const text = r.ok ? '' : await r.text();
    return { ok: r.ok, status: r.status, error: r.ok ? null : spError(text) };
  }
  async function get(suffix) {
    const r = await fetchWithRetry(apiUrl(suffix), {
      method: 'GET', headers: { 'Accept': 'application/json;odata=verbose' },
    });
    if (!r.ok) return { ok: false, status: r.status, error: spError(await r.text()) };
    return { ok: true, status: r.status, d: (await r.json()).d };
  }
  // An item POST's __metadata.type must be the LIST'S OWN entity type
  // (SP.Data.<MangledListName>ListItem), not the generic SP.Data.ListItem,
  // which the first version of this probe sent, earning two HTTP 400s that
  // it then reported as "Seeded two rows". deploy.js has always resolved
  // this properly; the probe did not.
  async function entityTypeFor(listTitle) {
    const r = await get(
      `web/lists/getbytitle('${odataName(listTitle)}')?$select=ListItemEntityTypeFullName`,
    );
    if (!r.ok) throw new Error(`could not resolve the item entity type: ${r.error}`);
    return r.d.ListItemEntityTypeFullName;
  }

  const listPath = `web/lists/getbytitle('${odataName(PROBE_LIST)}')`;
  let createdList = false;
  let viewId = null;
  let viewUrl = null;

  try {
    // === Setup: a list, a number column, two rows, and a view =============
    const made = await post('web/lists', {
      __metadata: { type: 'SP.List' },
      BaseTemplate: 100,
      Title: PROBE_LIST,
      Description: 'dbml-sharepoint aggregations probe. Safe to delete.',
    });
    if (!made.ok) {
      record('Q1', 'setup', 'ABORTED', `could not create the probe list: HTTP ${made.status} ${made.error}`);
      throw new Error('setup failed');
    }
    createdList = true;
    log('INFO', `Created '${PROBE_LIST}'.`);

    const field = await post(`${listPath}/fields`, {
      __metadata: { type: 'SP.FieldNumber' },
      FieldTypeKind: 9,
      Title: AGG_FIELD,
    });
    if (!field.ok) {
      record('Q1', 'setup', 'ABORTED', `could not add ${AGG_FIELD}: HTTP ${field.status} ${field.error}`);
      throw new Error('setup failed');
    }
    // Seed, and CHECK. The first version of this probe fired these two
    // posts, ignored their responses and logged "Seeded two rows"; both had
    // returned HTTP 400 and the list was empty for the manual step that
    // followed. A probe that reports an unchecked claim is worse than no
    // probe, so the seeded count is now a recorded question of its own and
    // is read back from the list rather than assumed.
    const itemType = await entityTypeFor(PROBE_LIST);
    const seedErrors = [];
    for (const amount of [10, 32]) {
      const seeded = await post(`${listPath}/items`, {
        __metadata: { type: itemType },
        Title: `probe ${amount}`,
        [AGG_FIELD]: amount,
      });
      if (!seeded.ok) seedErrors.push(`${amount}: HTTP ${seeded.status} ${seeded.error}`);
    }
    const counted = await get(`${listPath}/ItemCount`);
    const rowCount = counted.ok ? Number(counted.d.ItemCount) : NaN;
    record(
      'Q0',
      'two rows actually seeded, so the manual check has something to total',
      rowCount === 2 ? 'SEEDED' : 'FAILED',
      seedErrors.length
        ? `${seedErrors.join('; ')} (ItemCount=${rowCount})`
        : `ItemCount=${rowCount}; ${AGG_TYPE} of ${AGG_FIELD} should be 42`,
    );
    if (rowCount !== 2) {
      log('ERROR', 'The list is not correctly seeded. Q4 cannot be answered by looking at it. Fix this before trusting the verdict.');
    }

    const view = await post(`${listPath}/views`, {
      __metadata: { type: 'SP.View' },
      Title: 'Probe totals',
      ViewQuery: '',
      RowLimit: 30,
    });
    if (!view.ok) {
      record('Q1', 'setup', 'ABORTED', `could not create the view: HTTP ${view.status} ${view.error}`);
      throw new Error('setup failed');
    }
    const views = await get(`${listPath}/views?$select=Id,Title,ServerRelativeUrl,Aggregations,AggregationsStatus`);
    const probeView = (views.d?.results || []).find((v) => v.Title === 'Probe totals');
    viewId = probeView?.Id;
    viewUrl = probeView?.ServerRelativeUrl;
    // The Amount column must be IN the view or there is nothing to total.
    await post(`${listPath}/views('${viewId}')/viewfields/addviewfield('${odataName(AGG_FIELD)}')`);
    log('INFO', `View ready at ${window.location.origin}${viewUrl}`);

    const AGG_XML = `<FieldRef Name="${AGG_FIELD}" Type="${AGG_TYPE}"/>`;

    // === Q1: REST PATCH of the SP.View properties ========================
    const patched = await post(
      `${listPath}/views('${viewId}')`,
      { __metadata: { type: 'SP.View' }, Aggregations: AGG_XML, AggregationsStatus: 'On' },
      { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
    );
    record(
      'Q1',
      'REST PATCH of SP.View Aggregations/AggregationsStatus is accepted',
      patched.ok ? 'ACCEPTED' : 'REFUSED',
      patched.ok ? `HTTP ${patched.status}` : `HTTP ${patched.status} (${patched.error})`,
    );

    // === Q2: the SetViewXml path, guarded exactly as widths are ==========
    // Only attempted if the PATCH was refused: if PATCH works it is the
    // simpler mechanism and the one the deployer should use.
    let xmlWorked = false;
    if (!patched.ok) {
      const current = await post(`${listPath}/views('${viewId}')/renderashtml`);
      const xmlGet = await get(`${listPath}/views('${viewId}')?$select=ListViewXml`);
      const currentXml = xmlGet.d?.ListViewXml || '';
      if (!currentXml) {
        record('Q2', 'GetViewXml/SetViewXml carries an <Aggregations> block', 'NOT ESTABLISHED', 'could not read ListViewXml');
      } else {
        const block = `<Aggregations Value="On">${AGG_XML}</Aggregations>`;
        const strip = (xml) => xml.replace(/<Aggregations[\s\S]*?<\/Aggregations>/, '');
        const nextXml = currentXml.includes('<Aggregations')
          ? currentXml.replace(/<Aggregations[\s\S]*?<\/Aggregations>/, block)
          : currentXml.replace('</View>', `${block}</View>`);
        // Same guard the deployer's width splice uses: refuse if anything
        // outside the spliced region would change.
        if (strip(nextXml) !== strip(currentXml)) {
          record('Q2', 'GetViewXml/SetViewXml carries an <Aggregations> block', 'ABORTED', 'splice guard tripped: non-Aggregations content would change');
        } else {
          const set = await post(
            `${listPath}/views('${viewId}')`,
            { __metadata: { type: 'SP.View' }, ListViewXml: nextXml },
            { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
          );
          xmlWorked = set.ok;
          record(
            'Q2',
            'GetViewXml/SetViewXml carries an <Aggregations> block',
            set.ok ? 'ACCEPTED' : 'REFUSED',
            set.ok ? `HTTP ${set.status}` : `HTTP ${set.status} (${set.error})`,
          );
        }
      }
    } else {
      record('Q2', 'GetViewXml/SetViewXml carries an <Aggregations> block', 'NOT ATTEMPTED', 'the PATCH in Q1 was accepted, so the simpler mechanism wins');
    }

    // === Q3: read it back ================================================
    const after = await get(`${listPath}/views('${viewId}')?$select=Aggregations,AggregationsStatus`);
    const gotAgg = after.d?.Aggregations || '(null)';
    const gotStatus = after.d?.AggregationsStatus || '(null)';
    const matches = String(gotAgg).includes(AGG_FIELD) && String(gotAgg).includes(AGG_TYPE);
    record(
      'Q3',
      'the written property reads back unchanged',
      matches ? 'ROUND-TRIPPED' : 'MISMATCH',
      `Aggregations=${gotAgg} AggregationsStatus=${gotStatus}`,
    );

    record('Q4', 'a totals row actually RENDERS', 'MANUAL', `open ${window.location.origin}${viewUrl} and look for a total of 42 under ${AGG_FIELD}`);

    // === Q5/Q6: the naming question, and two columns at once =============
    //
    // WHY Q5 MATTERS MORE THAN THE REST. The sibling ColumnWidth property
    // binds by DISPLAY title. Internal names are accepted and silently
    // reset the widths (live finding, see jsgen). Aggregations is written
    // with INTERNAL names on the assumption that it does not behave the
    // same way, and the original run of this probe could not tell: it
    // created a field whose Title equalled its internal name, so the two
    // hypotheses were indistinguishable.
    //
    // Every template that ships totals uses display_names: auto, so every
    // totalled column's display title DIFFERS from its internal name,
    // i.e. all of them are in the case this probe never covered. If the
    // property binds by display title, the XML round-trips, both verify
    // halves pass, and no figure renders. That is the exact silent
    // failure this repository exists to prevent.
    const SECOND = 'SecondAmount';
    const second = await post(`${listPath}/fields`, {
      __metadata: { type: 'SP.FieldNumber' }, FieldTypeKind: 9, Title: SECOND,
    });
    if (second.ok) {
      // Rename the DISPLAY title, leaving the internal name as created,
      // the same create-then-rename trick deploy.js uses for every column.
      await post(
        `${listPath}/fields/getbyinternalnameortitle('${odataName(SECOND)}')`,
        { __metadata: { type: 'SP.FieldNumber' }, Title: 'Second Amount Display' },
        { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
      );
      await post(`${listPath}/views('${viewId}')/viewfields/addviewfield('${odataName(SECOND)}')`);

      // SEED THE SECOND COLUMN, because run 1 could not answer its own
      // question without it. The field is created after the rows, so both
      // were empty, and SharePoint renders NOTHING for an aggregation over
      // an empty column (no label, no zero). An operator looking for a
      // figure under "Second Amount Display" therefore saw exactly what a
      // FAILED BINDING would look like, and Q5 could only be settled by
      // hand-typing values into the list. That ambiguity is the probe's
      // fault; these two writes remove it.
      const secondValues = [1, 3];  // AVG 2, distinct from Amount's Sum 42
      const seededRows = await get(`${listPath}/items?$select=Id&$top=10`);
      const seededIds = (seededRows.ok && seededRows.d && seededRows.d.results
        ? seededRows.d.results.map((r) => r.Id) : []);
      for (let i = 0; i < seededIds.length && i < secondValues.length; i += 1) {
        await post(
          `${listPath}/items(${seededIds[i]})`,
          { __metadata: { type: itemType }, [SECOND]: secondValues[i] },
          { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
        );
      }
      const both = `<FieldRef Name="${AGG_FIELD}" Type="${AGG_TYPE}"/>`
        + `<FieldRef Name="${SECOND}" Type="AVG"/>`;
      const wrote = await post(
        `${listPath}/views('${viewId}')`,
        { __metadata: { type: 'SP.View' }, Aggregations: both, AggregationsStatus: 'On' },
        { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' },
      );
      const back = await get(`${listPath}/views('${viewId}')?$select=Aggregations`);
      record(
        'Q5',
        'Aggregations binds by INTERNAL name, not display title',
        wrote.ok ? 'MANUAL' : 'WRITE REFUSED',
        wrote.ok
          ? `wrote Name="${SECOND}" while its DISPLAY title is "Second Amount Display"; `
            + `readback ${back.ok ? JSON.stringify(back.d.Aggregations) : '(unreadable)'}. `
            + `OPEN THE VIEW: a figure under "Second Amount Display" means INTERNAL names bind `
            + `(what the tool assumes). No figure under it means DISPLAY titles bind, and every `
            + `shipped totals view is silently empty.`
          : `HTTP ${wrote.status} (${wrote.error})`,
      );
      record(
        'Q6',
        'two totalled columns both render, in declaration order',
        wrote.ok ? 'MANUAL' : 'NOT REACHED',
        wrote.ok
          ? 'the same view now declares two aggregations; confirm BOTH figures appear, and that '
            + 'the readback above preserved declaration order. The deployer compares the string '
            + 'exactly, so a reordered readback would drift on every redeploy'
          : 'the two-column write was refused',
      );
    } else {
      record('Q5', 'Aggregations binds by INTERNAL name, not display title', 'NOT ESTABLISHED', `could not add ${SECOND}: HTTP ${second.status} ${second.error}`);
      record('Q6', 'two totalled columns both render, in declaration order', 'NOT ESTABLISHED', 'the second column could not be created');
    }

    // === Verdict =========================================================
    const mechanism = patched.ok ? 'patch' : (xmlWorked ? 'setviewxml' : 'none');
    console.table(results.map(({ id, question, observed, detail }) => ({ id, question, observed, detail })));
    log(
      'VERDICT',
      `seeded=${rowCount === 2 ? 'ok' : 'FAILED'} mechanism=${mechanism} `
      + `readback=${matches ? 'ok' : 'mismatch'} rendered=<fill in after looking>`,
    );
    if (rowCount !== 2) {
      log('ERROR', 'seeded=FAILED. An empty list shows no totals row whether the feature works or not, so `rendered=no` from this run would mean nothing. Fix the seeding and re-run before reporting.');
    }
    log('INFO', 'Paste the VERDICT line back, with rendered= set to yes or no.');
    return { seeded: rowCount, mechanism, readback: matches ? 'ok' : 'mismatch', viewUrl, results };
  } finally {
    if (createdList && CLEANUP_AT_END) {
      const gone = await post(listPath, undefined, { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' });
      if (gone.ok) {
        log('INFO', `Deleted '${PROBE_LIST}'.`);
      } else {
        log('ERROR', `COULD NOT DELETE '${PROBE_LIST}' (HTTP ${gone.status} ${gone.error}). Delete it by hand: ${window.location.origin}${WEB}/Lists/${encodeURIComponent(PROBE_LIST)}`);
      }
    } else if (createdList) {
      log('INFO', `Left '${PROBE_LIST}' in place for the manual step. Re-run with CLEANUP_AT_END = true to remove it.`);
    }
  }
})();
