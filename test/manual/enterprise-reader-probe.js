/**
 * dbml-sharepoint PROBE: THE ENTERPRISE READER TIER
 *
 * STATUS: RUN TWICE, on ONE live site (2026-08-11 and 2026-08-12). The
 * second run was pasted BY THE REPORTING SERVICE ACCOUNT ITSELF with the
 * write flag ON, so GROUP B IS ANSWERED. A2, A4 and C2-C3 are still open,
 * and the A2/A3 CONTRADICTION recorded below is UNEXPLAINED. Read that
 * note before quoting anything from run 2's group A.
 *
 *   RUN 1:    2026-08-11, probe revision 0e565ebc, pasted into ONE
 *             Microsoft 365 group-connected Team Site (NOT a publishing
 *             site) by an OPERATOR who owns it. Read-only: ALLOW_WRITES
 *             was off, so group B did not run.
 *   RUN 2:    2026-08-12, the same site, pasted BY THE REPORTING SERVICE
 *             ACCOUNT (the reader itself, not an operator) with
 *             ALLOW_WRITES ON, so group B did run.
 *   FINDINGS: run 1 answered A1-A5 in full and C1; its C2 and C3 WERE AN
 *             INVALID EXPERIMENT and have been rewritten since (see GROUP
 *             C below for what was wrong and what replaced it); B1-B4
 *             were blocked by the write flag. Run 2 answered B1-B4 in
 *             full, repeated A5, and CORROBORATED A3 from a second
 *             identity with its control open; A2, A4 and C1-C3 were not
 *             answered.
 *
 * ---- WHAT THE 2026-08-11 RUN MEASURED ---------------------------------
 *   A1  The nominated service account resolved READ-ONLY out of the site's
 *       existing users, to a claims LoginName. Prerequisite met, so A3 and
 *       A4 are about that account.
 *   A2  THE CONTROL HELD. getusereffectivepermissions answered for the
 *       account pasting the probe and decoded to a full permission set,
 *       so a narrow A3 is a fact about the READER, not about an endpoint
 *       the caller cannot use.
 *   A3  THE HEADLINE. At WEB scope the reader account held EXACTLY:
 *           ViewFormPages, Open, BrowseUserInfo, UseClientIntegration,
 *           UseRemoteAPIs
 *       That is precisely Learn's documented `Limited Access` set, and
 *       UseRemoteAPIs IS PRESENT. The lockdown-mode failure mode this
 *       whole feature was hedged against DID NOT OCCUR ON THIS SITE.
 *   A4  At LIST scope, on a list that read back
 *       HasUniqueRoleAssignments=true, the reader held:
 *           ViewListItems, OpenItems, ViewVersions, ViewFormPages, Open,
 *           ViewPages, CreateSSCSite, BrowseUserInfo,
 *           UseClientIntegration, UseRemoteAPIs, CreateAlerts
 *       That is the built-in `Read`. So the declared grant DOES reach
 *       through broken inheritance (which is the mechanism the entire
 *       tier depends on), measured, on one site.
 *   A5  The ViewFormPagesLockDown feature GUID was ABSENT from the 15 site
 *       collection features that came back. Consistent with A3. The REST
 *       spelling used to read them is undocumented, so this CORROBORATES
 *       A3; it does not independently prove lockdown mode is off.
 *
 * ---- WHAT THE 2026-08-12 RUN MEASURED ---------------------------------
 * WHO RAN IT FIRST, because it changes what every row below means. The
 * REPORTING SERVICE ACCOUNT pasted this itself. That account is NOT A
 * MEMBER OF THE SITE: it holds list-scoped Read through the readers group
 * plus the Limited Access SharePoint derives at web scope, and nothing
 * else. It could nonetheless sign in, open the browser console, run this
 * script and call REST. That is an observation about this account on this
 * site, not a rule about SharePoint.
 *
 *   B1+B2  THE IMPORTANT ONE, and it corrects a comment that shipped.
 *          `web/ensureuser` for BOTH display names ("Everyone" AND
 *          "Everyone except external users") resolved to ONE AND THE SAME
 *          principal: a `c:0-.f|rolemanager|spo-grid-all-users/<tenantid>`
 *          LoginName, Title "Everyone except external users",
 *          PRINCIPALTYPE 4 (SecurityGroup), Email empty. Two of the three
 *          needles `_reader_enrolment.js.j2` ships matched it
 *          (`spo-grid-all-users` and `c:0-.f|rolemanager`).
 *          So on this tenant the strict `PrincipalType !== 1` check at
 *          step 2 of that template refuses both on its own, and the step-3
 *          comment which said they "resolve as ordinary users, so the type
 *          check above does not catch them" was written from plausibility
 *          and is FALSE. It has been corrected. NOTHING WAS WEAKENED: one
 *          tenant is one data point, another tenant may type these
 *          differently, and the needles stay as defence in depth.
 *   B3+B4  "All Authenticated Users" and "All Forms Users" were BOTH
 *          REFUSED (HTTP 400, SPException, "The specified user … could
 *          not be found") under odata=nometadata and odata=verbose
 *          alike. They are not resolvable BY DISPLAY NAME on this tenant.
 *          That NARROWS the KNOWN LIMIT at step 3 of the deploy template;
 *          it does not close it. A display-name refusal is not proof that
 *          no login-name encoding exists, and another tenant may resolve
 *          them.
 *   A3     Decoded to EXACTLY the five permissions run 1 recorded:
 *          ViewFormPages, Open, BrowseUserInfo, UseClientIntegration,
 *          UseRemoteAPIs. Measured from a SECOND IDENTITY, which makes it
 *          CORROBORATION AND NOTHING MORE. A2 was open on this run, so
 *          run 1's claim is NOT upgraded by it. See the next section.
 *   A5     The ViewFormPagesLockDown feature GUID was again ABSENT from
 *          the 15 site collection features that came back.
 *   OPEN   A2 returned HTTP 403 UnauthorizedAccessException. A4 returned
 *          HTTP 403. C1-C3 never ran: the configured GROUP_NAME did not
 *          exist on that site, so the prerequisite was unmet and
 *          SERVER-DRIVEN PAGING ON A GROUP'S MEMBERS IS STILL UNMEASURED
 *          after two runs.
 *
 * ---- THE A2/A3 CONTRADICTION, AND IT IS UNEXPLAINED --------------------
 * In run 2 the caller and the reader were THE SAME ACCOUNT, so A2 and A3
 * issued the same request. A2 called getusereffectivepermissions at WEB
 * scope for that login and got HTTP 403 UnauthorizedAccessException. A3
 * then called THE SAME ENDPOINT at the same scope for the same login and
 * got HTTP 200 with a decodable permission set. 403 then 200, seconds
 * apart, nothing else changed.
 *
 * NOTHING ESTABLISHED HERE EXPLAINS THAT. The most parsimonious
 * HYPOTHESIS is that the first permissions call of a run can fail
 * transiently (some auth or context warm-up) and A2 simply goes first.
 * THAT IS A HYPOTHESIS, NOT A FINDING: its only support is the ordering.
 * A RE-RUN IS WHAT WOULD SETTLE IT. If A2 holds next time with nothing
 * else changed, the ordering story gains weight; if A2 fails again while
 * A3 succeeds again, this endpoint is doing something the probe does not
 * understand and no reader row from it can be trusted until it does.
 *
 * NO RETRY HAS BEEN ADDED, deliberately. A retry would turn this into a
 * green run and destroy the only evidence that it happens; showing it is
 * the probe's job. What HAS changed is that A3 and A4 now INHERIT A2's
 * outcome. While the control is open they record NOT ESTABLISHED
 * (control open) and carry their decoded bits as supporting detail only.
 * In run 2 A3 recorded OBSERVED while its own control was refused, and a
 * row that contradicts its own control must not read as a finding. That
 * was a defect in this probe and it is fixed.
 *
 * ---- THE SCOPE LIMIT, AND IT IS NOT A FORMALITY ------------------------
 * ONE SITE, and it is a Microsoft 365 GROUP-CONNECTED TEAM SITE. Lockdown
 * mode is ON BY DEFAULT FOR PUBLISHING SITES, so A3 DOES NOT CLEAR THAT
 * CASE. The configuration most likely to break this feature is the one
 * configuration this run did not sample. One site is one data point. State
 * both of those every time this run is cited.
 *
 * ---- STILL OPEN AFTER BOTH RUNS ---------------------------------------
 *   - A2, THE CONTROL, and with it A3 and A4 as findings. See the
 *     contradiction note above. This is the first thing a third run should
 *     answer, because every reader row depends on it.
 *   - WHETHER THE POWER BI CONNECTOR ITSELF WORKS. At web scope the
 *     account holds no ViewPages and no ViewListItems (see A3), so
 *     enumerating `_api/web/lists` (which the SharePoint Online List
 *     connector does once you hand it a site URL) is UNPROVEN either way.
 *     Run 1 said only signing in AS the service account could answer this
 *     and that no revision of this file could; RUN 2 WAS THE SERVICE
 *     ACCOUNT, so that is no longer true. What is still missing is the
 *     question itself: this probe asks nothing about `web/lists`, and it
 *     would have to be added.
 *   - The same A3/A4 measurement on a PUBLISHING site.
 *   - A4 at list scope from the reader's own session: refused HTTP 403 in
 *     run 2 and only ever answered in run 1's operator session.
 *   - The two tenant-wide claim encodings Learn does not publish. B3 and
 *     B4 have narrowed this (both display names were refused HTTP 400 on
 *     this tenant), but a refusal by display name is not proof that no
 *     encoding exists.
 *   - Server-driven paging on a site group's members: see GROUP C. Run 2
 *     could not reach it at all. C2 IS NOT BEING ASKED FOR AGAIN. It
 *     needs a site group with more members than the server page size, and
 *     adding one to a site's permissions would expose many real people to
 *     development work. That objection is accepted and settled. The
 *     nearest measurable substitute is S11/S12 in
 *     search-discovery-probe.js.j2, which pages the LIST-ITEMS collection
 *     against an existing 6,000-row fixture: it can establish whether this
 *     tenant's REST emits server-driven paging AT ALL, but it does not
 *     settle the site-group case, which stays unmeasured.
 *
 * ---- C2/C3 WERE AN INVALID EXPERIMENT; THEY HAVE BEEN REWRITTEN --------
 * As run on 2026-08-11, C2 asked `web/sitegroups(id)/users?$top=1` against
 * a two-member group, found no `d.__next`, and C3 concluded from that that
 * both deploy templates "would have read only the first page of this
 * group." THAT CONCLUSION DOES NOT FOLLOW. It is removed rather than left
 * in the file where somebody could quote it.
 *
 * `$top` is CLIENT-DRIVEN paging. Learn, "PageSize, Top and MaxTop": "In
 * client-driven paging, we request the server to return the specified
 * number of results. There is no `nextLink` that is returned." And Learn,
 * "Server-driven Paging in ASP.NET Core OData 8": if `$top` is below the
 * page size, the server returns that many items "with no next link." So an
 * absent `__next` at `$top=1` says NOTHING about whether a collection
 * larger than the server page size returns one. The experiment could not
 * have produced a `__next` even if the endpoint pages perfectly.
 *
 * What the run DID establish about paging is C1, and that stands: both
 * members arrived in ONE page at `$top=5000`, in the `d.results` shape
 * both deploy templates assume.
 *
 * A VALID experiment needs a group with MORE MEMBERS THAN THE SERVER PAGE
 * SIZE, read WITHOUT a `$top` below that (the deploy sends `$top=5000`).
 * That is what C2 and C3 do now. SharePoint does not document the server
 * page size for this collection, so C2 assumes no number: it reads with no
 * `$top` at all and compares the row count against C1's, which lets it
 * tell "the server paged" from "the group fitted in one page". The second
 * is recorded as NOT ESTABLISHED, never as a finding about the endpoint.
 *
 * WHY THIS EXISTS. The `feat/enterprise-reader` branch gives each of the 31
 * shipped families an empty '<XX> Enterprise Readers' site group holding the
 * built-in `Read` on every list, and `build --enterprise-reader <upn>` puts
 * one named service account into it. Because those lists BREAK INHERITANCE,
 * the account ends up with `Read` at LIST scope and only SharePoint's derived
 * `Limited Access` at WEB scope. The customer then points Power BI Pro at
 * many such sites using an Entra user service account.
 *
 * Three things that whole design rests on have never been measured on a
 * tenant. This probe measures them and asserts none of them.
 *
 * ---- WHAT IS ALREADY ESTABLISHED ON MICROSOFT LEARN --------------------
 * These are cited, not re-derived, and they are the reason the questions
 * below are the RIGHT questions:
 *
 *   - `_api/web/getusereffectivepermissions(@user)?@user='<claims login>'`
 *     returns an SP.BasePermissions, a High/Low bitfield. The same method
 *     exists on a list. (SharePoint .NET Server, CSOM, JSOM and REST API
 *     index; Web/List.GetUserEffectivePermissions.)
 *   - Built-in `Read` carries `Use Remote Interfaces`. `Restricted Read`
 *     does not. ("Permission levels in SharePoint".)
 *   - `Limited Access` carries `Use Remote Interfaces` BY DEFAULT, and
 *     LOCKDOWN MODE STRIPS IT. Lockdown mode ("Limited-access user
 *     permission lockdown mode") is ON BY DEFAULT FOR PUBLISHING SITES.
 *     (Same page, "Lockdown mode" section, and the troubleshooting article
 *     "Creating a team site from SharePoint Home doesn't finish", which
 *     names UseRemoteAPI as the permission lockdown mode removes.)
 *
 * So the feature is safe iff the account really does hold UseRemoteAPIs at
 * web scope and ViewListItems at list scope, on the sites it will be
 * pointed at. That is question group A, and nobody has looked.
 *
 * ---- WHAT IT ASKS -----------------------------------------------------
 *
 * GROUP A, the decisive one. Read-only; runs with the write flag OFF.
 *   A1  Does the nominated account resolve READ-ONLY in this site's User
 *       Information List, and to what claims LoginName? A prerequisite for
 *       A2/A3, recorded as its own row so an unresolvable account reads as
 *       "not asked" rather than as a permission finding.
 *   A2  CONTROL: does getusereffectivepermissions work AT ALL for the
 *       account pasting this, at web scope? Without it, a blank A3 is
 *       indistinguishable from an endpoint the caller cannot use. It
 *       OBSERVES the decode; it asserts nothing about which bits come back.
 *       ITS OUTCOME IS THREADED INTO A3 AND A4. While this row is open
 *       those two record NOT ESTABLISHED (control open) whatever they read.
 *   A3  The reader's effective permissions at WEB scope, decoded into
 *       named permissions. HEADLINE OBSERVATION: is UseRemoteAPIs present?
 *       A finding only if A2 held; otherwise supporting detail.
 *   A4  The reader's effective permissions at LIST scope for the nominated
 *       list, decoded. HEADLINE OBSERVATION: is ViewListItems present?
 *       A finding only if A2 held; otherwise supporting detail.
 *   A5  Is the 'Limited-access user permission lockdown mode' site
 *       collection feature active? See the LOCKDOWN MODE note below. The
 *       feature GUID is documented, the REST spelling for reading it is
 *       not, so this row is best-effort and says so in its own evidence.
 *
 * GROUP B, the two claim encodings this branch deliberately did not guess.
 *   WRITE-GATED. Default OFF.
 *   `_reader_enrolment.js.j2` step 3 refuses tenant-wide claims by matching
 *   three login-name needles, which cover two of the four claims Learn
 *   names ("Grant the Everyone claim to external users in Microsoft 365"):
 *   Everyone, Everyone except external users, All Authenticated Users, All
 *   Forms Users. Learn publishes no login-name encoding for the last two,
 *   so they were left out ON PURPOSE with a dated KNOWN LIMIT comment
 *   rather than invented from plausibility.
 *   B1..B4 call `web/ensureuser` for each of the four DISPLAY NAMES and
 *   RECORD the LoginName, Title and PrincipalType that come back, or the
 *   exact failure if it refuses. Each row also reports whether the three
 *   needles the deploy template ships today would have matched what came
 *   back, an observation about the existing guard, not a requirement.
 *
 *   NOTHING HERE IS A PREDICTION. If a claim resolves, that is the answer.
 *   If it refuses, that is equally the answer, and the status and body are
 *   printed verbatim either way.
 *
 *   ANSWERED 2026-08-12, on one tenant (see the run-2 section above). B1
 *   and B2 collapsed to a single PrincipalType 4 principal that two
 *   needles matched; B3 and B4 were refused HTTP 400. Re-run these on any
 *   other tenant you can get at: that is the only thing that turns one
 *   data point into a pattern.
 *
 * GROUP C, a paging assumption this branch inherited but never measured.
 *   Read-only.
 *   `_security_principals.js.j2` and `_reader_enrolment.js.j2` both page a
 *   site group's members with `web/sitegroups(<id>)/users?$top=5000` and
 *   follow `d.__next`. Nobody has confirmed that endpoint pages with
 *   `__next` at all, or in that shape. C1 reads the whole membership the
 *   way the deploy does and records how many members came back, which is
 *   the yardstick for C2; C2 reads the SAME collection with NO `$top`, so
 *   only the SERVER decides where a page ends, and reports whether it
 *   returned a short page and a `d.__next`; C3 follows that `__next` and
 *   reports whether a further page of different members arrives.
 *
 *   C2 CAN ONLY ANSWER ON A GROUP BIGGER THAN THE SERVER PAGE SIZE, which
 *   SharePoint does not document. On a small group it records NOT
 *   ESTABLISHED and says so. See the C2/C3 note in STATUS above for the
 *   invalid version of this experiment and why an absent `__next` under a
 *   small `$top` proves nothing.
 *
 *   DO NOT CREATE OR ENROL A LARGE GROUP TO SATISFY THIS. That has been
 *   declined and the reason stands: it would expose many real people to
 *   development work. Run C against whatever groups already exist, accept
 *   NOT ESTABLISHED when they are small, and read S11/S12 in
 *   search-discovery-probe.js.j2 for the part of the question that CAN be
 *   measured, server-driven paging on a list-items collection, against a
 *   fixture that already exists.
 *
 * ---- WHAT THE WRITE FLAG GUARDS ---------------------------------------
 * ALLOW_WRITES gates GROUP B AND NOTHING ELSE. Groups A and C are pure
 * reads and run with it off.
 *
 * `web/ensureuser` IS A WRITE. It materialises the principal in this
 * site's User Information List (hidden list 'User Information List'), which
 * is why the deploy template takes a form digest for it. IT GRANTS
 * NOTHING. It does not add anybody to a group, does not create a role
 * assignment, and does not change any permission. The side effect is that
 * up to four claim principals may afterwards appear in this site's people
 * picker and in `web/siteusers`. On a disposable probe site that is
 * harmless; on a real site it is untidy and hard to fully undo, so RUN
 * GROUP B ON A SITE YOU ARE WILLING TO THROW AWAY.
 *
 * ---- LOCKDOWN MODE: WHAT IS NOT ESTABLISHED ---------------------------
 * Learn documents the FEATURE (site collection feature "Limited-access
 * user permission lockdown mode") and documents its ID as
 * 7C637B23-06C4-472D-9A9A-7C175762C5C4 under the display name
 * ViewFormPagesLockDown ("SPMT supported SharePoint site features").
 * Learn also documents `Site.Features` as a CSOM property.
 *
 * WHAT LEARN DOES NOT PUBLISH, as far as this author could find, is a REST
 * spelling for reading that feature collection. A5 therefore TRIES the
 * OData mapping of Site.Features and reports exactly what came back,
 * including a non-2xx status and its body. A failed read is recorded as
 * NOT ESTABLISHED, never as "lockdown is off". Absence of evidence here
 * is not evidence of absence, and getting that backwards would report the
 * dangerous configuration as the safe one. If A5 comes back NOT
 * ESTABLISHED, read lockdown mode by eye instead: Site settings > Site
 * collection features, and say which it was.
 *
 * ---- HOW TO RUN -------------------------------------------------------
 *   1. Edit the four constants under CONFIGURATION below. Every one ships
 *      as an obvious placeholder and the probe refuses the group that
 *      needs it until it has been changed. There is deliberately no site
 *      URL: the probe reads the site it is pasted into.
 *   2. Open the target SharePoint site as somebody who can read
 *      permissions there (a site owner or site collection administrator).
 *      F12 -> Console -> paste -> Enter. It prints its plan and stops.
 *   3. Set CONFIRMED = true and paste again. That runs GROUPS A AND C,
 *      which write nothing.
 *   4. ONLY on a disposable site, and only if you want group B answered:
 *      set ALLOW_WRITES = true as well and paste once more.
 *   5. Copy the whole RESULTS block back verbatim, including the raw
 *      payload lines. The transcript stays OUT of this repository. Quote
 *      findings into the STATUS block at the top of this file instead.
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

  // Printed FIRST, before any gate: a stale clipboard and a fix that did
  // not work produce identical transcripts otherwise.
  log('INFO', 'probe revision 279fe884. Quote this when reporting results.');

  // ---- CONFIGURATION ---------------------------------------------------
  // All three are obvious placeholders. Each group refuses to run against
  // an unedited one and says which constant to set, so an unedited paste
  // never produces a row that looks like a measurement.
  //
  // NO SITE URL, deliberately. See the harness.

  // The service account the customer would point Power BI at. A UPN.
  const READER_UPN = 'reporting-service-account@contoso.com';

  // Optional: the claims LoginName, if you already have it (it looks like
  // 'i:0#.f|membership|someone@contoso.com'). Leave empty and A1 resolves
  // it read-only from the site's existing users instead.
  const READER_LOGIN_NAME = '';

  // A list this bundle provisioned, one that BROKE INHERITANCE and carries
  // the Enterprise Readers group's Read grant. That is the whole point of
  // A4: web scope and list scope are expected to differ.
  const LIST_TITLE = 'CHANGE ME - a list this bundle provisioned';

  // The site group to measure paging against, for group C. It does NOT have
  // to be the readers group. Use a group THAT ALREADY EXISTS. C2 can only
  // answer on a group with MORE members than SharePoint's undocumented
  // server page size, and it records NOT ESTABLISHED (not a finding) when
  // the whole membership arrives in one page.
  //
  // DO NOT ADD PEOPLE OR A LARGE GROUP TO A SITE TO MAKE C2 ANSWERABLE.
  // That was declined and the reason stands; see GROUP C in the docblock.
  const GROUP_NAME = 'CHANGE ME - the largest site group that already exists';

  const PLACEHOLDER = /CHANGE ME|reporting-service-account@contoso\.com/;

  // ---- Registration ----------------------------------------------------
  expect('A1', 'Does the nominated account resolve read-only, and to what claims LoginName?');
  expect('A2', 'CONTROL: does getusereffectivepermissions work at all for the account running this?');
  expect('A3', 'The reader\'s effective permissions at WEB scope: is UseRemoteAPIs present?');
  expect('A4', 'The reader\'s effective permissions at LIST scope: is ViewListItems present?');
  expect('A5', 'Is the limited-access user permission lockdown mode site collection feature active?');
  expect('B1', 'What does web/ensureuser return for the display name "Everyone"?');
  expect('B2', 'What does web/ensureuser return for "Everyone except external users"?');
  expect('B3', 'What does web/ensureuser return for "All Authenticated Users"?');
  expect('B4', 'What does web/ensureuser return for "All Forms Users"?');
  expect('C1', 'How many members does web/sitegroups(id)/users return at $top=5000, as the deploy reads it?');
  expect('C2', 'With NO $top, does web/sitegroups(id)/users return a short page and a d.__next?');
  expect('C3', 'Does following d.__next return a further page of DIFFERENT members?');

  if (!CONFIRMED) {
    log('INFO', 'Would READ, writing nothing: the nominated account\'s effective');
    log('INFO', 'permissions at web scope and at list scope, decoded into named');
    log('INFO', 'permissions; the lockdown-mode site collection feature; and one');
    log('INFO', 'site group\'s membership twice, once the way the deploy reads it,');
    log('INFO', 'once with no $top at all, to see whether the SERVER pages it.');
    log('INFO', 'Would ADDITIONALLY, only with ALLOW_WRITES = true, call');
    log('INFO', 'web/ensureuser for four tenant-wide claim display names. That is a');
    log('INFO', 'WRITE (it materialises each principal in this site\'s User');
    log('INFO', 'Information List), but it GRANTS NOTHING: no group membership, no');
    log('INFO', 'role assignment, no permission change. Use a disposable site.');
    log('INFO', 'Nothing has been read or written. Set CONFIRMED = true.');
    return;
  }

  // Everything measured sits inside this try; see the `finally` at the end.
  try {
  // ---- SP.BasePermissions decoding -------------------------------------
  // MIRRORED, NOT INVENTED. Every value below is copied character for
  // character from src/dbml_sharepoint/analysis/permissions.py's
  // BASE_PERMISSIONS, which the deployer already uses to EMIT permission
  // levels, so the probe reads permissions with the same bit positions the
  // tool writes them with, and a divergence would be a bug in one file
  // rather than two independent guesses. Kept in the same order as the
  // Python dict so the two can be diffed by eye.
  //
  // BigInt, because these are 64-bit and JavaScript's bitwise operators are
  // 32-bit: `0x4000000000000000 & x` with plain numbers is silently wrong.
  const BASE_PERMISSIONS = [
    ['ViewListItems',             0x0000000000000001n],
    ['AddListItems',              0x0000000000000002n],
    ['EditListItems',             0x0000000000000004n],
    ['DeleteListItems',           0x0000000000000008n],
    ['ApproveItems',              0x0000000000000010n],
    ['OpenItems',                 0x0000000000000020n],
    ['ViewVersions',              0x0000000000000040n],
    ['DeleteVersions',            0x0000000000000080n],
    ['CancelCheckout',            0x0000000000000100n],
    ['ManagePersonalViews',       0x0000000000000200n],
    ['ManageLists',               0x0000000000000800n],
    ['ViewFormPages',             0x0000000000001000n],
    ['AnonymousSearchAccessList', 0x0000000000002000n],
    ['Open',                      0x0000000000010000n],
    ['ViewPages',                 0x0000000000020000n],
    ['AddAndCustomizePages',      0x0000000000040000n],
    ['ApplyThemeAndBorder',       0x0000000000080000n],
    ['ApplyStyleSheets',          0x0000000000100000n],
    ['ViewUsageData',             0x0000000000200000n],
    ['CreateSSCSite',             0x0000000000400000n],
    ['ManageSubwebs',             0x0000000000800000n],
    ['CreateGroups',              0x0000000001000000n],
    ['ManagePermissions',         0x0000000002000000n],
    ['BrowseDirectories',         0x0000000004000000n],
    ['BrowseUserInfo',            0x0000000008000000n],
    ['AddDelPrivateWebParts',     0x0000000010000000n],
    ['UpdatePersonalWebParts',    0x0000000020000000n],
    ['ManageWeb',                 0x0000000040000000n],
    ['UseClientIntegration',      0x0000001000000000n],
    ['UseRemoteAPIs',             0x0000002000000000n],
    ['ManageAlerts',              0x0000004000000000n],
    ['CreateAlerts',              0x0000008000000000n],
    ['EditMyUserInfo',            0x0000010000000000n],
    ['EnumeratePermissions',      0x4000000000000000n],
  ];
  // EmptyMask (0) and FullMask are in the Python table too and are left out
  // of the loop above on purpose: a zero mask matches every value, so
  // including it would put 'EmptyMask' in the decoded list of literally
  // every result, and FullMask is reported separately below.
  const FULL_MASK = 0x7FFFFFFFFFFFFFFFn;

  // High/Low may arrive as strings or as numbers depending on the OData
  // flavour, so both are coerced through String first. Each is at most
  // 2^32, so no precision is lost on the way in.
  const toMask = (high, low) => (BigInt(String(high)) << 32n) | BigInt(String(low));
  const decodeMask = (mask) => BASE_PERMISSIONS
    .filter(([, bit]) => (mask & bit) === bit)
    .map(([name]) => name);

  // The RESPONSE SHAPE of getusereffectivepermissions is itself an
  // observation, not something this probe knows. Verbose wraps the result
  // under d.<MethodName>; nometadata is flat; some endpoints wrap in
  // 'value'. Every shape this can find is tried and the one that WORKED is
  // reported alongside the numbers, and the raw payload is always printed
  // so a reader can see a shape this ladder does not know about.
  const extractHighLow = (payload) => {
    const candidates = [
      ['flat', payload],
      ['value', payload && payload.value],
      ['d', payload && payload.d],
      ['d.GetUserEffectivePermissions', payload && payload.d && payload.d.GetUserEffectivePermissions],
      ['GetUserEffectivePermissions', payload && payload.GetUserEffectivePermissions],
    ];
    for (const [shape, node] of candidates) {
      if (node && node.High != null && node.Low != null) {
        return { shape, high: node.High, low: node.Low };
      }
    }
    return null;
  };

  // ---- Read helpers ----------------------------------------------------
  // The harness's spGet sends odata=nometadata. Group C has to reproduce
  // what the DEPLOY TEMPLATES send, and `d.__next` is a VERBOSE construct,
  // so verbose gets its own helper rather than being bolted onto spGet.
  // Both keep the harness's contract: `body` is the parsed payload whether
  // or not the request succeeded, so only `ok` says the call worked.
  const spGetVerbose = async (path) => {
    const res = await fetch(`${WEB}/_api/${path}`, {
      headers: { Accept: 'application/json;odata=verbose' },
    });
    const text = await res.text();
    let parsed = null;
    try { parsed = JSON.parse(text); } catch { /* SharePoint sent plain text */ }
    return { ok: res.ok, status: res.status, body: parsed, text };
  };

  // OData string literal: double the single quotes, then percent-encode.
  // encodeURIComponent leaves "'" alone, so the doubling survives, and it
  // DOES encode '#' and '|', which every claims login name contains and
  // which would otherwise truncate the URL at the fragment.
  const odataLiteral = (value) => encodeURIComponent(String(value).replace(/'/g, "''"));

  // Every failure in this probe is reported AS ITS STATUS WITH ITS BODY.
  // A previous probe here collapsed an unexplained HTTP error into
  // "REFUSED" and produced a transcript that read like a measured platform
  // verdict; the ampersand it blamed had nothing to do with it. isRefusal
  // (harness) already carves out identity and throttling, and even a real
  // refusal keeps its status and body attached.
  const httpDetail = (r) => `HTTP ${r.status}: ${String(r.text == null ? JSON.stringify(r.body) : r.text).slice(0, 500)}`;
  const failureOutcome = (r) => (isRefusal(r.status) ? `REFUSED (HTTP ${r.status})` : `NOT ESTABLISHED (HTTP ${r.status})`);

  // ================= GROUP A: read-only ================================

  // ---- A1: resolve the account WITHOUT writing --------------------------
  // `ensureuser` would resolve anything, but it is a write, and group A
  // has to be answerable with the write flag off. So this reads the users
  // the site already knows about. If the account has never touched the
  // site it will not be there, and that is a PREREQUISITE NOT MET, not a
  // permission finding: the difference is the whole reason this is its own
  // row.
  let readerLogin = READER_LOGIN_NAME;
  {
    if (readerLogin) {
      record('A1', 'Does the nominated account resolve read-only, and to what claims LoginName?',
             'OBSERVED: supplied by the operator',
             `READER_LOGIN_NAME was set, so no lookup was needed: ${JSON.stringify(readerLogin)}`);
    } else if (PLACEHOLDER.test(READER_UPN)) {
      record('A1', 'Does the nominated account resolve read-only, and to what claims LoginName?',
             'NOT ESTABLISHED (prerequisite)',
             'READER_UPN is still the placeholder. Set it to the service account UPN and paste again.');
    } else {
      // Two reads, cheapest first. Both are recorded so a reader can see
      // WHICH one answered. The $filter form is the one worth reusing in
      // the deployer if it works, and its working is an observation.
      const wanted = READER_UPN.toLowerCase();
      const filtered = await spGet(
        `web/siteusers?$select=Id,LoginName,Title,Email,PrincipalType&$filter=Email eq '${odataLiteral(READER_UPN)}'`);
      let via = "$filter=Email eq '...'";
      let match = null;
      if (!readFailed(filtered) && Array.isArray(filtered.body.value)) {
        match = filtered.body.value.find((u) => String(u.Email || '').toLowerCase() === wanted) || null;
      }
      let enumerated = null;
      if (!match) {
        // Fall back to the whole collection and match the same two ways
        // _reader_enrolment.js.j2 step 4 does: the mailbox Email, and the
        // UPN the claims LoginName ends in. They are routinely different
        // for the same account.
        via = 'enumerating web/siteusers';
        enumerated = await spGet('web/siteusers?$select=Id,LoginName,Title,Email,PrincipalType&$top=5000');
        const upnPart = (name) => {
          const parts = String(name).split('|');
          return parts[parts.length - 1];
        };
        const rows = (!readFailed(enumerated) && Array.isArray(enumerated.body.value)) ? enumerated.body.value : [];
        match = rows.find((u) => [u.Email, upnPart(u.LoginName)]
          .map((v) => String(v == null ? '' : v).toLowerCase())
          .includes(wanted)) || null;
      }
      if (match && typeof match.LoginName === 'string' && match.LoginName !== '') {
        readerLogin = match.LoginName;
        record('A1', 'Does the nominated account resolve read-only, and to what claims LoginName?',
               'OBSERVED: resolved',
               `via ${via}: LoginName ${JSON.stringify(match.LoginName)}, Title `
               + `${JSON.stringify(match.Title)}, PrincipalType ${JSON.stringify(match.PrincipalType)}`);
      } else {
        const detail = enumerated
          ? `$filter read ${readFailed(filtered) ? httpDetail(filtered) : 'succeeded but matched nobody'}; `
            + `enumeration ${readFailed(enumerated) ? httpDetail(enumerated) : 'succeeded but matched nobody'}`
          : `$filter read ${readFailed(filtered) ? httpDetail(filtered) : 'succeeded but matched nobody'}`;
        record('A1', 'Does the nominated account resolve read-only, and to what claims LoginName?',
               'NOT ESTABLISHED (prerequisite)',
               `${detail}. The account may simply never have been added to this site. That is not a `
               + 'permission finding. Add it (or run the deploy with --enterprise-reader), or paste its '
               + 'claims login into READER_LOGIN_NAME, and run again.');
      }
    }
  }

  // ---- Shared effective-permissions reader ------------------------------
  // depends on: the scope path and the login name, both fixed by the
  // caller. observes: the HTTP status, the response SHAPE, and the bits.
  // Nothing here compares an observed bit against an expectation. The
  // outcome string names what was seen and the decode is printed in full.
  const readEffectivePermissions = async (scopePath, login) => {
    const path = `${scopePath}/getusereffectivepermissions(@user)?@user='${odataLiteral(login)}'`;
    const res = await spGet(path);
    if (readFailed(res)) return { ok: false, res, path };
    const found = extractHighLow(res.body);
    if (!found) return { ok: false, res, path, unknownShape: true };
    const mask = toMask(found.high, found.low);
    return {
      ok: true, res, path, mask, shape: found.shape,
      high: String(found.high), low: String(found.low),
      names: decodeMask(mask), full: mask === FULL_MASK,
    };
  };
  const describePerms = (r) => `High=${r.high} Low=${r.low} (response shape: ${r.shape}); `
    + `decoded: ${r.names.length ? r.names.join(', ') : '(no named permission bit set)'}`
    + `${r.full ? '; this is FullMask' : ''}; raw ${JSON.stringify(r.res.body).slice(0, 400)}`;

  // ---- A2: the control --------------------------------------------------
  // Answered for whoever is pasting, not for the reader. Its job is to
  // separate "the endpoint does not work for this caller" from "the reader
  // holds nothing", two states that produce identical-looking empty rows.
  // It reports WHAT it decoded and does not require any particular bit:
  // asserting over an observed value is how a control kills its own
  // experiment the moment it starts working.
  //
  // ITS OUTCOME IS THREADED INTO A3 AND A4, and that is not tidiness.
  // 2026-08-12: A2 came back HTTP 403 and A3 (the same endpoint, the same
  // web scope, the same login, because the service account was running its
  // own probe) came back HTTP 200 and RECORDED ITSELF AS OBSERVED. A2's
  // evidence said in words that A3 could not be read as a finding while
  // the control was open, and the row above it did exactly that anyway. A
  // caveat in prose beside a row that reads OBSERVED loses to the row. So
  // the control's verdict is now carried in a variable and A3/A4 downgrade
  // themselves to NOT ESTABLISHED (control open) when it did not hold.
  let controlHeld = false;
  let controlWhy = 'A2 did not run';
  let callerLogin = null;
  {
    const me = await spGet('web/currentuser?$select=LoginName,Title');
    if (readFailed(me) || typeof me.body.LoginName !== 'string') {
      controlWhy = `web/currentuser could not be read: ${httpDetail(me)}`;
      record('A2', 'CONTROL: does getusereffectivepermissions work at all for the account running this?',
             'NOT ESTABLISHED', `could not read web/currentuser: ${httpDetail(me)}`);
    } else {
      callerLogin = me.body.LoginName;
      const mine = await readEffectivePermissions('web', callerLogin);
      if (!mine.ok) {
        controlWhy = mine.unknownShape
          ? 'A2 got HTTP 200 but no High/Low it could recognise'
          : `A2 got ${httpDetail(mine.res)}`;
        record('A2', 'CONTROL: does getusereffectivepermissions work at all for the account running this?',
               mine.unknownShape ? 'NOT ESTABLISHED (unrecognised response shape)' : failureOutcome(mine.res),
               `${mine.unknownShape ? 'HTTP 200 but no High/Low found in ' + JSON.stringify(mine.res.body).slice(0, 400)
                                    : httpDetail(mine.res)}. `
               + 'A3 and A4 below cannot be read as findings about the reader while this row is open. The '
               + 'endpoint has not been shown to work for the caller at all. They will say so themselves.');
      } else {
        controlHeld = true;
        record('A2', 'CONTROL: does getusereffectivepermissions work at all for the account running this?',
               'OBSERVED: the endpoint answered',
               `for the CALLER ${JSON.stringify(me.body.Title)}: ${describePerms(mine)}`);
      }
    }
  }

  // What a reader row says when the control did not hold. The bits are
  // still printed (throwing away a successful read would be worse), but
  // they are labelled as supporting detail for a re-run to compare
  // against, never as an answer.
  //
  // The same-login case is called out explicitly, because that is the one
  // that happened and it is the strongest form of the contradiction: two
  // calls differing in nothing but their order, answering differently.
  // See THE A2/A3 CONTRADICTION in the docblock. It is not retried here on
  // purpose. A retry would hide the very thing worth seeing.
  const sameLoginAsControl = (login) =>
    typeof callerLogin === 'string' && String(login).toLowerCase() === callerLogin.toLowerCase();
  const controlOpenEvidence = (login, headline, detail) =>
    `NOT a finding about the reader: the control (A2) has not been shown to work for this caller. `
    + `${controlWhy}. Supporting detail only: ${headline}; ${detail}.`
    + (sameLoginAsControl(login)
      ? ' NOTE THE CONTRADICTION: this row used THE SAME LOGIN and THE SAME endpoint as the control, so'
        + ' the two calls differ in nothing but their order, and they did not answer the same way. That is'
        + ' unexplained. Re-run before treating any of it as measured.'
      : '');

  // ---- A3: WEB scope ----------------------------------------------------
  {
    const question = 'The reader\'s effective permissions at WEB scope: is UseRemoteAPIs present?';
    if (!readerLogin) {
      record('A3', question, 'NOT ESTABLISHED (prerequisite)', 'the account did not resolve; see A1');
    } else {
      const r = await readEffectivePermissions('web', readerLogin);
      if (!r.ok) {
        record('A3', question,
               r.unknownShape ? 'NOT ESTABLISHED (unrecognised response shape)' : failureOutcome(r.res),
               r.unknownShape
                 ? `HTTP 200 but no High/Low found in ${JSON.stringify(r.res.body).slice(0, 400)}`
                 : httpDetail(r.res));
      } else {
        const has = r.names.includes('UseRemoteAPIs');
        const headline = `UseRemoteAPIs ${has ? 'PRESENT' : 'ABSENT'}`;
        if (!controlHeld) {
          record('A3', question, 'NOT ESTABLISHED (control open)',
                 controlOpenEvidence(readerLogin, headline, describePerms(r)));
        } else {
          record('A3', question, `OBSERVED: ${headline}`, describePerms(r));
        }
      }
    }
  }

  // ---- A4: LIST scope ---------------------------------------------------
  {
    const question = 'The reader\'s effective permissions at LIST scope: is ViewListItems present?';
    if (PLACEHOLDER.test(LIST_TITLE)) {
      record('A4', question, 'NOT ESTABLISHED (prerequisite)',
             'LIST_TITLE is still the placeholder. Set it to a list this bundle provisioned.');
    } else if (!readerLogin) {
      record('A4', question, 'NOT ESTABLISHED (prerequisite)', 'the account did not resolve; see A1');
    } else {
      // Confirm the list exists FIRST, so a typo reads as a typo. Without
      // this, a 404 for a misspelled title and a genuine refusal look the
      // same in the transcript, and the wrong one is much more alarming.
      const listRes = await spGet(
        `web/lists/getbytitle('${odataLiteral(LIST_TITLE)}')?$select=Title,HasUniqueRoleAssignments`);
      if (readFailed(listRes)) {
        record('A4', question, 'NOT ESTABLISHED (prerequisite)',
               `no list titled ${JSON.stringify(LIST_TITLE)} could be read here: ${httpDetail(listRes)}`);
      } else {
        // Recorded, not required: whether the list actually broke
        // inheritance is the thing that makes list scope differ from web
        // scope, so it belongs in the evidence, but a list that inherits
        // is still a real answer about a real list.
        const unique = listRes.body.HasUniqueRoleAssignments;
        const r = await readEffectivePermissions(
          `web/lists/getbytitle('${odataLiteral(LIST_TITLE)}')`, readerLogin);
        if (!r.ok) {
          record('A4', question,
                 r.unknownShape ? 'NOT ESTABLISHED (unrecognised response shape)' : failureOutcome(r.res),
                 r.unknownShape
                   ? `HTTP 200 but no High/Low found in ${JSON.stringify(r.res.body).slice(0, 400)}`
                   : httpDetail(r.res));
        } else {
          const has = r.names.includes('ViewListItems');
          const headline = `ViewListItems ${has ? 'PRESENT' : 'ABSENT'}`;
          const detail = `list ${JSON.stringify(LIST_TITLE)}, `
            + `HasUniqueRoleAssignments=${JSON.stringify(unique)}; ${describePerms(r)}`;
          if (!controlHeld) {
            record('A4', question, 'NOT ESTABLISHED (control open)',
                   controlOpenEvidence(readerLogin, headline, detail));
          } else {
            record('A4', question, `OBSERVED: ${headline}`, detail);
          }
        }
      }
    }
  }

  // ---- A5: lockdown mode, best-effort -----------------------------------
  // See the LOCKDOWN MODE note in the docblock. The feature ID is
  // documented; the REST spelling for reading the collection is not, so
  // the read itself is part of what is being observed. A failure here is
  // recorded as NOT ESTABLISHED and never as "lockdown is off".
  {
    const question = 'Is the limited-access user permission lockdown mode site collection feature active?';
    const LOCKDOWN_ID = '7c637b23-06c4-472d-9a9a-7c175762c5c4';
    const res = await spGet('site/features?$select=DefinitionId&$top=500');
    const rows = (!readFailed(res) && Array.isArray(res.body && res.body.value)) ? res.body.value : null;
    if (rows === null) {
      record('A5', question, 'NOT ESTABLISHED',
             `${readFailed(res) ? httpDetail(res)
                                : `HTTP ${res.status} but no feature array in ${JSON.stringify(res.body).slice(0, 300)}`}`
             + '. This is NOT evidence that lockdown mode is off. The REST spelling used here is not '
             + 'documented on Learn, and the caller may simply not be able to read site features. Check by '
             + 'eye instead: Site settings > Site collection features > "Limited-access user permission '
             + 'lockdown mode", and report Active or Not Active.');
    } else {
      const ids = rows
        .map((f) => String((f && (f.DefinitionId || (f.DefinitionId === 0 ? 0 : ''))) || '').toLowerCase())
        .filter((id) => id !== '');
      const active = ids.includes(LOCKDOWN_ID);
      record('A5', question,
             `OBSERVED: lockdown mode ${active ? 'ACTIVE' : 'NOT LISTED AS ACTIVE'}`,
             `read ${rows.length} site collection feature(s); looked for ${LOCKDOWN_ID} `
             + '(ViewFormPagesLockDown, per Learn\'s SPMT supported site features table). '
             + `${active ? 'It is present.' : 'It is absent from the list that came back.'} `
             + 'Confirm by eye in Site settings > Site collection features before relying on this. The '
             + 'endpoint spelling is undocumented and the collection may be filtered for the caller.');
    }
  }

  // ================= GROUP C: read-only ================================
  // Placed before group B so that a run with ALLOW_WRITES off still ends
  // having answered everything it can.
  {
    const C1_Q = 'How many members does web/sitegroups(id)/users return at $top=5000, as the deploy reads it?';
    const C2_Q = 'With NO $top, does web/sitegroups(id)/users return a short page and a d.__next?';
    const C3_Q = 'Does following d.__next return a further page of DIFFERENT members?';
    if (PLACEHOLDER.test(GROUP_NAME)) {
      const why = 'GROUP_NAME is still the placeholder. Set it to a site group, and see C2: only a group '
        + 'with MORE members than the server page size can answer C2 and C3.';
      record('C1', C1_Q, 'NOT ESTABLISHED (prerequisite)', why);
      record('C2', C2_Q, 'NOT ESTABLISHED (prerequisite)', why);
      record('C3', C3_Q, 'NOT ESTABLISHED (prerequisite)', why);
    } else {
      const grpRes = await spGetVerbose(`web/sitegroups/getbyname('${odataLiteral(GROUP_NAME)}')?$select=Id`);
      const groupId = (grpRes.ok && grpRes.body && grpRes.body.d && grpRes.body.d.Id);
      if (!Number.isInteger(groupId)) {
        const why = `could not resolve site group ${JSON.stringify(GROUP_NAME)}: ${httpDetail(grpRes)}`;
        record('C1', C1_Q, 'NOT ESTABLISHED (prerequisite)', why);
        record('C2', C2_Q, 'NOT ESTABLISHED (prerequisite)', why);
        record('C3', C3_Q, 'NOT ESTABLISHED (prerequisite)', why);
      } else {
        // C1: read exactly the way the deploy templates read it. Its row
        // count is the YARDSTICK C2 needs: without it, C2 cannot tell a
        // page the server truncated from the whole membership arriving at
        // once, and those two look identical in the response.
        const allRes = await spGetVerbose(`web/sitegroups(${groupId})/users?$select=Id,LoginName,Title&$top=5000`);
        const all = (allRes.ok && allRes.body && allRes.body.d && Array.isArray(allRes.body.d.results))
          ? allRes.body.d.results : null;
        if (all === null) {
          const why = `membership read failed or was not in the d.results shape: ${httpDetail(allRes)}`;
          // NOT `failureOutcome(allRes)`. That returns `REFUSED (HTTP n)` for
          // any non-2xx outside 401/403/408/429, and `report()` counts a row
          // as OPEN only when its outcome starts with NOT ESTABLISHED or
          // SHORT -- so a 404 or a 500 here was summarised as C1 ANSWERED
          // while it had measured no member count at all.
          //
          // A refusal IS a genuine answer in group B, where the question is
          // what `web/ensureuser` returns for a display name: being refused
          // is the finding. C1 asks for a ROW COUNT and a `d.__next`. A read
          // that returned neither answered nothing, whatever the status.
          record('C1', C1_Q, `NOT ESTABLISHED (HTTP ${allRes.status})`, why);
          record('C2', C2_Q, 'NOT ESTABLISHED (prerequisite)', why);
          record('C3', C3_Q, 'NOT ESTABLISHED (prerequisite)', why);
        } else {
          const allNext = allRes.body.d.__next;
          record('C1', C1_Q, `OBSERVED: ${all.length} member(s)`,
                 `${all.length} member(s) came back at $top=5000, in the d.results shape both deploy `
                 + `templates assume; d.__next on that read: ${JSON.stringify(
                   typeof allNext === 'string' && allNext !== '')}. This is the count C2 compares against.`);

          // C2: THE SAME REQUEST WITH NO $top AT ALL, which is the only
          // form that can answer the question. $top is CLIENT-driven
          // paging: Learn ("PageSize, Top and MaxTop") says there is no
          // nextLink returned for it, and ("Server-driven Paging in
          // ASP.NET Core OData 8") that a $top below the page size returns
          // that many items with no next link. So a small $top cannot
          // produce a __next however well the endpoint pages, and the
          // earlier $top=1 version of this question was measuring nothing.
          // With no $top the SERVER alone decides where the page ends.
          const firstRes = await spGetVerbose(`web/sitegroups(${groupId})/users?$select=Id,LoginName,Title`);
          const firstPage = (firstRes.ok && firstRes.body && firstRes.body.d
                             && Array.isArray(firstRes.body.d.results)) ? firstRes.body.d.results : null;
          if (firstPage === null) {
            const why = `the no-$top read failed or was not in the d.results shape: ${httpDetail(firstRes)}`;
            // Same reason as C1 above: C2 asks whether the SERVER hands back
            // a `d.__next`, so a read that did not come back cannot answer it.
            record('C2', C2_Q, `NOT ESTABLISHED (HTTP ${firstRes.status})`, why);
            record('C3', C3_Q, 'NOT ESTABLISHED (prerequisite)', why);
          } else if (firstPage.length >= all.length) {
            // The whole membership fitted in one server page, so the
            // server never had to page and an absent __next is expected.
            // THIS IS NOT A FINDING ABOUT THE ENDPOINT. Recording it as
            // one is exactly the mistake the 2026-08-11 run made.
            const why = `the no-$top read returned ${firstPage.length} row(s), which is every member C1 `
              + `counted (${all.length}). The server never had to page, so whether d.__next came back `
              + `(${JSON.stringify(typeof firstRes.body.d.__next === 'string'
                                   && firstRes.body.d.__next !== '')}) says NOTHING about whether this `
              + 'endpoint pages. A valid answer needs a site group with MORE members than the server page '
              + 'size, which SharePoint does not document. DO NOT BUILD ONE: enrolling a large group to '
              + 'satisfy this has been declined, because it would expose many real people to development '
              + 'work. This row stays open. For the part of the question that CAN be measured, see '
              + 'S11/S12 in search-discovery-probe.js.j2, which pages a list-items collection against an '
              + 'existing 6,000-row fixture. Do NOT read this row as "the endpoint does not page".';
            record('C2', C2_Q, 'NOT ESTABLISHED (group not larger than the server page size)', why);
            record('C3', C3_Q, 'NOT ESTABLISHED (prerequisite)', why);
          } else {
            const next = firstRes.body.d.__next;
            // The site host is stripped before printing. __next is an
            // absolute URL, so quoting it raw would put the tenant into the
            // transcript. This repository has leaked one twice. Whether it
            // WAS absolute is still recorded, because that is the part the
            // deploy's `membersUrl = json.d.__next` depends on.
            // Case-INSENSITIVE, and deliberately not a RegExp: SharePoint may
            // echo the host or path in a different case than
            // _spPageContextInfo reports, and a case-sensitive split would
            // leave the tenant in the transcript. Building a RegExp from a URL
            // would mean escaping it correctly, which is one more thing to get
            // wrong in the code whose whole job is not leaking the tenant.
            const redactWeb = (s) => {
              const hay = s.toLowerCase();
              const needle = WEB.toLowerCase();
              let out = '';
              let from = 0;
              for (;;) {
                const at = hay.indexOf(needle, from);
                if (at === -1) { return out + s.slice(from); }
                out += s.slice(from, at) + '<site>';
                from = at + needle.length;
              }
            };
            const redacted = typeof next === 'string' ? redactWeb(next) : next;
            const present = typeof next === 'string' && next !== '';
            // The server truncated: fewer rows than C1 counted, with no
            // $top asked for. So this run CAN discriminate, and both
            // outcomes are real answers about the endpoint.
            record('C2', C2_Q,
                   `OBSERVED: server paged at ${firstPage.length} of ${all.length}; d.__next `
                   + `${present ? 'PRESENT' : 'ABSENT'}`,
                   `no $top was sent; ${firstPage.length} row(s) came back against the ${all.length} C1 `
                   + `counted, so the SERVER ended the page. typeof d.__next = ${typeof next}; absolute `
                   + `(starts with this site's URL): ${present ? String(next.startsWith(WEB)) : 'n/a'}; `
                   + `value with the site host redacted: ${JSON.stringify(redacted)}`
                   + `${present ? '' : '. A truncated page with NO continuation is a finding about the '
                       + 'deploy templates: both stop when __next is absent, so they would silently read '
                       + 'a PARTIAL membership of a group this size.'}`);

            // C3: follow it. The deploy passes d.__next straight back to
            // fetch, so this does too, rather than rebuilding a URL.
            if (!present) {
              record('C3', C3_Q, 'NOT ESTABLISHED (prerequisite)',
                     'there was no d.__next to follow; see C2, which records what that means.');
            } else {
              const res2 = await fetch(next, { headers: { Accept: 'application/json;odata=verbose' } });
              const text2 = await res2.text();
              let body2 = null;
              try { body2 = JSON.parse(text2); } catch { /* not JSON */ }
              const page2 = (res2.ok && body2 && body2.d && Array.isArray(body2.d.results)) ? body2.d.results : null;
              if (page2 === null) {
                record('C3', C3_Q,
                       isRefusal(res2.status) ? `REFUSED (HTTP ${res2.status})` : `NOT ESTABLISHED (HTTP ${res2.status})`,
                       `following d.__next returned HTTP ${res2.status}: ${text2.slice(0, 400)}`);
              } else {
                // Whole-set comparison, not first-row: a second page that
                // repeats the first would still have a different first Id
                // if the order changed, and that would read as success.
                const seen = new Set(firstPage.map((u) => u.Id));
                const fresh = page2.filter((u) => !seen.has(u.Id)).length;
                record('C3', C3_Q,
                       `OBSERVED: page 2 held ${page2.length} row(s), ${fresh} of them NOT on page 1`,
                       `page 1 held ${firstPage.length} row(s); page 2 held ${page2.length}, of which `
                       + `${fresh} carried an Id page 1 did not; page 2 __next present: `
                       + `${JSON.stringify(typeof body2.d.__next === 'string' && body2.d.__next !== '')}; `
                       + `running total ${seen.size + fresh} against the ${all.length} C1 counted`);
              }
            }
          }
        }
      }
    }
  }

  // ================= GROUP B: WRITE-GATED ==============================
  const CLAIM_NAMES = [
    ['B1', 'Everyone'],
    ['B2', 'Everyone except external users'],
    ['B3', 'All Authenticated Users'],
    ['B4', 'All Forms Users'],
  ];
  const claimQuestion = (display) => `What does web/ensureuser return for the display name "${display}"?`;

  if (!ALLOW_WRITES) {
    for (const [id, display] of CLAIM_NAMES) {
      record(id, claimQuestion(display), 'NOT ESTABLISHED (write flag off)',
             'ensureuser materialises the principal in this site\'s User Information List, so it is a '
             + 'write and is gated. It GRANTS NOTHING: no group membership, no role assignment, no '
             + 'permission change. Set ALLOW_WRITES = true on a DISPOSABLE site to answer this.');
    }
  } else {
    log('INFO', 'ALLOW_WRITES is on: calling web/ensureuser for four claim display names.');
    log('INFO', 'This materialises each resolved principal in this site\'s User Information');
    log('INFO', 'List. It grants nothing, but the principals may afterwards appear in this');
    log('INFO', 'site\'s people picker and in web/siteusers.');

    // The three needles _reader_enrolment.js.j2 step 3 ships TODAY. Copied
    // verbatim so each row can report whether the guard as it stands would
    // have caught what came back. That is an OBSERVATION about the guard,
    // not a requirement on the answer. A claim that no needle matches is
    // exactly the finding this probe exists to produce.
    const SHIPPED_NEEDLES = ['spo-grid-all-users', 'c:0(.s|true', 'c:0-.f|rolemanager'];

    // Two flavours, because the deploy sends verbose and the harness sends
    // nometadata, and which of them ensureuser accepts is not something to
    // assume: a nometadata rejection reported as "the claim was refused"
    // would be a fabricated platform verdict. Whichever answered is named
    // in the evidence.
    const ensureUser = async (logonName) => {
      const select = 'web/ensureuser?$select=Id,LoginName,Title,Email,PrincipalType';
      const digestA = await getDigest();
      const flat = await spPost(select, { logonName }, digestA);
      if (flat.ok) return { attempt: 'odata=nometadata', res: flat, node: flat.body };
      const digestB = await getDigest();
      const verbose = await spPost(select, { logonName }, digestB, {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
      });
      if (verbose.ok) {
        return { attempt: 'odata=verbose (nometadata was refused first)', res: verbose,
                 node: verbose.body && verbose.body.d };
      }
      return { attempt: 'both', res: flat, fallback: verbose, node: null };
    };

    for (const [id, display] of CLAIM_NAMES) {
      const question = claimQuestion(display);
      const out = await ensureUser(display);
      if (!out.node) {
        const extra = out.fallback
          ? ` The odata=verbose retry also failed: ${httpDetail(out.fallback)}.`
          : '';
        record(id, question, failureOutcome(out.res),
               `${httpDetail(out.res)}.${extra} Reported as the status it actually returned. This row `
               + 'says what the platform did, not whether the claim "should" exist.');
        continue;
      }
      const login = String(out.node.LoginName == null ? '' : out.node.LoginName);
      const matched = SHIPPED_NEEDLES.filter((n) => login.toLowerCase().includes(n));
      record(id, question, 'OBSERVED: resolved',
             `via ${out.attempt}: LoginName ${JSON.stringify(out.node.LoginName)}, Title `
             + `${JSON.stringify(out.node.Title)}, PrincipalType ${JSON.stringify(out.node.PrincipalType)}, `
             + `Email ${JSON.stringify(out.node.Email)}, Id ${JSON.stringify(out.node.Id)}. `
             + `The three needles _reader_enrolment.js.j2 ships today would ${matched.length
                 ? `MATCH this (${matched.join(', ')})` : 'NOT match this'}.`);
    }
  }

  } finally {
    // ---- Report --------------------------------------------------------
    // In a `finally`, so a throw anywhere in A/B/C still prints every
    // question answered before it. A probe that loses its whole transcript
    // to one failed call has wasted the operator's paste.
    report();

  // The trailer is INSIDE the `finally` too, and that is the point. It used
  // to sit after the try/finally, so a throw printed the whole RESULTS block
  // -- site host, account UPN, real principal ids -- and then skipped the
  // line telling the operator not to commit it, along with questions 2 to 5.
  // The transcript is at its most sensitive exactly when the run failed, and
  // that was the one path with no warning attached. A tenant URL has leaked
  // from this repository twice.
  console.log('\n============ WHAT TO SEND BACK ============');
  console.log('1. The whole RESULTS block above, verbatim, including the evidence');
  console.log('   lines, the decoded permission names are the finding, not the');
  console.log('   PRESENT/ABSENT word beside them.');
  console.log('2. Whether the site you pasted this into is a PUBLISHING site.');
  console.log('   Lockdown mode is on by default for those, and that is exactly');
  console.log('   what would strip UseRemoteAPIs from Limited Access.');
  console.log('   answer: ______________________________________');
  console.log('3. If A5 came back NOT ESTABLISHED: Site settings > Site collection');
  console.log('   features > "Limited-access user permission lockdown mode":');
  console.log('   Active or Not Active?');
  console.log('   answer: ______________________________________');
  console.log('4. Whether ALLOW_WRITES was on for this run (group B is blank if not).');
  console.log('   answer: ______________________________________');
  console.log('5. WHICH ACCOUNT you pasted this as: an operator/site owner, or the');
  console.log('   reporting service account itself. It changes what group A means,');
  console.log('   and on 2026-08-12 it was the reader, which is how A2 and A3 came');
  console.log('   to disagree about one endpoint. Say which.');
  console.log('   answer: ______________________________________');
  console.log('\nDo NOT commit this transcript. It carries the site host, the account');
  console.log('UPN and real principal ids. Quote the findings into the STATUS block');
  console.log('at the top of enterprise-reader-probe.js.j2 instead.');
  console.log('===========================================');
  }
})();
