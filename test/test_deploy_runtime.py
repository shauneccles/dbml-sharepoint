# test/test_deploy_runtime.py
"""Execute the generated deploy.js against a mock SharePoint.

The golden-file test proves deploy.js does not CHANGE; it cannot prove it
RUNS. A whole class of defect lives in that gap (a caller that omits a
key another function requires, a comparison against `undefined`, a
sentinel that reads as a real value). One such bug shipped in the golden
fixture and was asserted as correct: the synthetic Title patch carried
none of the declared-formula keys, so every field reconcile treated it as
managed and aborted the phase on every list, on every run.

Node is required; the test skips without it rather than failing, since it
is not a dependency of the package.
"""

import json
import re
import subprocess
import tempfile
import textwrap
from pathlib import Path
from typing import Any
from urllib.parse import quote

import pytest
from _builders import ID_PK, table
from _node import NODE
from _node import run_node as _run
from _packs import DEFAULT_PREFIX, blocks, entities, pack
from _paths import FIXTURES

from dbml_sharepoint.analysis.phases import phase_number as pn


def _deploy_js() -> str:
    from dbml_sharepoint.generators.jsgen import generate_deploy_js
    from dbml_sharepoint.model.mapping_loader import load_mapping
    from dbml_sharepoint.model.parser import parse_dbml
    from dbml_sharepoint.model.release import load_release

    return generate_deploy_js(
        schema=parse_dbml(FIXTURES / "simple.dbml"),
        bundle=load_mapping(FIXTURES / "sharepoint-mapping.yaml"),
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="simple.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )


# A SharePoint that answers every read as an EMPTY, healthy list: no fields
# exist, no formulas are set. That is the state of a brand-new site, and the
# state in which the shipped bug threw.
_HARNESS = textwrap.dedent("""
    const calls = [];
    globalThis.window = { location: { origin: 'https://example.sharepoint.com' } };
    globalThis._spPageContextInfo = {
      webServerRelativeUrl: '/sites/test',
      userLoginName: 'probe@example.com',
      userId: 1,
    };
    const body = (url) => {
      if (url.includes('contextinfo')) {
        return { d: { GetContextWebInformation: {
          FormDigestValue: 'digest', FormDigestTimeoutSeconds: 1800 } } };
      }
      if (url.toLowerCase().includes('effectivebasepermissions')) {
        return { d: { EffectiveBasePermissions: { High: 4294967295, Low: 4294967295 } } };
      }
      if (url.includes('ClientValidationFormula') || url.includes('ValidationFormula')) {
        return { d: {
          ClientValidationFormula: null, ClientValidationMessage: null,
          ValidationFormula: null, ValidationMessage: null } };
      }
      return { d: { results: [] } };
    };
    globalThis.fetch = async (url, opts = {}) => {
      // body is null, never absent: JSON.stringify drops an undefined key,
      // and the Python side reads c['body'] unconditionally.
      calls.push({ url: String(url), method: opts.method || 'GET',
                   body: opts.body === undefined ? null : opts.body });
      return {
        ok: true, status: 200,
        headers: { get: () => null },
        json: async () => body(String(url)),
        text: async () => JSON.stringify(body(String(url))),
      };
    };
    globalThis.__calls = calls;
""")


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_deploy_js_runs_without_throwing() -> None:
    """The generated script must reach a summary against a healthy site.

    It need not succeed at provisioning (the mock is too thin for that),
    but a thrown exception or an abort carrying schema errors means the
    script is broken for every operator on every site.
    """
    script = _HARNESS + "\n" + _deploy_js().replace(
        "})();", "}))().then(r => console.log('__RESULT__' + JSON.stringify(r)))"
        ).replace("(async () => {", "((async () => {", 1)
    output = _run(script)
    result_line = next(
        (ln for ln in output.splitlines() if ln.startswith("__RESULT__")), None,
    )
    assert result_line is not None, f"deploy.js did not return a summary:\n{output[-3000:]}"
    summary = json.loads(result_line.removeprefix("__RESULT__"))
    # The mock is deliberately thin, so shape probes legitimately complain.
    # What must never appear is a formula error: that is the phase-aborting
    # failure the synthetic Title patch produced on every list, every run,
    # and it is invisible to a golden-file comparison.
    formula_errors = [
        err for err in (summary.get("errors") or [])
        if "ValidationFormula" in str(err) or "ValidationMessage" in str(err)
    ]
    assert not formula_errors, f"deploy.js aborted on the declared formulas: {formula_errors}"


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_the_builtin_title_column_is_never_sent_a_formula() -> None:
    """Title is not a declared field, so the tool does not own its
    formulas. The shipped bug MERGEd an empty ClientValidationMessage onto
    it before aborting, an unrequested write to a built-in column."""
    script = _HARNESS + "\n" + _deploy_js().replace(
        "})();", "}))().then(() => console.log('__CALLS__' + JSON.stringify(globalThis.__calls)))",
        ).replace("(async () => {", "((async () => {", 1)
    line = next(
        (ln for ln in _run(script).splitlines() if ln.startswith("__CALLS__")),
        None,
    )
    assert line is not None, "harness produced no call log"
    calls = json.loads(line.removeprefix("__CALLS__"))
    title_writes = [
        c for c in calls
        if c["method"] == "POST" and c["body"] and "ValidationMessage" in c["body"]
        and "'Title'" in c["url"]
    ]
    assert not title_writes, (
        f"wrote formula properties to the built-in Title column: {title_writes}"
    )


# An ADOPTED site: the lists already exist and the built-in Title is
# SEALED. The harness above answers every field probe as absent, so the
# adoption path (the one where declared shapes are actually compared)
# has never executed in a test. That is the gap the synthetic-Title bug
# shipped through.
_ADOPTED_HARNESS = textwrap.dedent(r"""
    const calls = [];
    globalThis.window = { location: { origin: 'https://example.sharepoint.com' } };
    globalThis._spPageContextInfo = {
      webServerRelativeUrl: '/sites/test',
      userLoginName: 'probe@example.com',
      userId: 1,
    };
    // Per-list Description state, mutated by MERGEs exactly as SharePoint
    // would, so the list probe reads back what was actually written and a
    // run can never satisfy its own read-back. Default '' is the honest
    // pre-marker state: a list provisioned before descriptions were written,
    // or one an owner blanked, which is exactly what the reconcile repairs.
    // Both constants are rewritten by _run_adopted_deploy.
    const LIST_DESCRIPTIONS = {};
    const IGNORE_DESCRIPTION_WRITES = false;
    const listDescription = (listTitle) => (
      LIST_DESCRIPTIONS[listTitle] == null ? '' : LIST_DESCRIPTIONS[listTitle]
    );
    // Per-group Description and paginated membership, keyed by group Title.
    // A name with no entry keeps the prior fixed shape (Description 'Test
    // group.', no members), so every existing test is unaffected. Rewritten
    // by _group_gate_deploy for the adoption-gate tests.
    const GROUP_DESCRIPTIONS = {};
    const GROUP_MEMBER_PAGES = {};
    // The site-group enumeration (web/sitegroups?$select=Title) that decides
    // whether a declared group reads as pre-existing or absent. Empty by
    // default, matching every other test's "brand-new site" fiction for
    // groups: the ADOPT branch below is otherwise unreachable, because the
    // enumeration fast path answers every declared group 404 before the
    // by-name probe this override targets ever runs.
    const KNOWN_GROUP_NAMES = [];
    const groupDescription = (name) => (
      GROUP_DESCRIPTIONS[name] == null ? 'Test group.' : GROUP_DESCRIPTIONS[name]
    );
    const groupMemberPages = (name) => GROUP_MEMBER_PAGES[name] || [[]];
    const groupNameOf = (url) => {
      const raw = (url.match(/sitegroups\/getbyname\('(.*?)'\)/) || [])[1];
      return raw == null ? raw : decodeURIComponent(raw).replace(/''/g, "'");
    };
    // Task 7 read-back verification. A group's write endpoints (create and
    // MERGE) mutate GROUP_STATE the same way the list write above mutates
    // LIST_DESCRIPTIONS, so verifyGroupSettings reads back what the mock
    // actually stored rather than a shape fixed in advance.
    // GROUP_DROP_FIELD_ON_WRITE names one property the tenant accepts but
    // never stores, modelling a write SharePoint 200s and discards.
    // GROUP_COERCE_AUTO_ACCEPT models the measured tenant behaviour
    // (test/manual/group-description-probe.js, G9/G10, 2026-08-13/14):
    // AutoAcceptRequestToJoinLeave is forced false whenever the written
    // AllowRequestToJoinLeave is false, regardless of what was sent for
    // AutoAccept itself.
    const GROUP_DROP_FIELD_ON_WRITE = null;
    const GROUP_COERCE_AUTO_ACCEPT = false;
    const GROUP_SETTINGS_KEYS = ['Description', 'AllowMembersEditMembership',
      'AllowRequestToJoinLeave', 'AutoAcceptRequestToJoinLeave',
      'OnlyAllowMembersViewMembership'];
    const GROUP_STATE = {};
    const groupState = (name) => (GROUP_STATE[name] ||= {
      Description: groupDescription(name), AllowMembersEditMembership: false,
      AllowRequestToJoinLeave: false, AutoAcceptRequestToJoinLeave: false,
      OnlyAllowMembersViewMembership: false,
    });
    // Existence for the by-name GET and its /users sub-resource: known from
    // the enumeration (case-insensitive, matching SharePoint's own group-name
    // resolution) or already written into GROUP_STATE by a create/MERGE this
    // run performed. Checked with hasOwnProperty rather than through
    // groupState() itself, whose `||=` would auto-vivify an absent name into
    // "existing" the instant it is asked about -- which is the exact hole
    // that let a read ahead of the create that makes it possible pass.
    const KNOWN_GROUP_NAME_SET = new Set(KNOWN_GROUP_NAMES.map((n) => String(n).toLowerCase()));
    const groupIsKnown = (name) => (
      KNOWN_GROUP_NAME_SET.has(String(name).toLowerCase())
      || Object.prototype.hasOwnProperty.call(GROUP_STATE, name)
    );
    // Group Id, historically fixed at 9 for every name -- harmless while the
    // owner probe below answered the same fixed value regardless of Id. A
    // test exercising TWO groups' owner state independently (one group
    // adopted, its declared owner_group a second, absent custom group) needs
    // them to resolve to DIFFERENT Ids, so this is now a per-name map,
    // defaulting every unconfigured name to the old fixed 9 so every
    // existing test is unaffected.
    const GROUP_IDS = {};
    const groupId = (name) => (GROUP_IDS[name] != null ? GROUP_IDS[name] : 9);
    // Current owner per governed-group Id (`web/sitegroups(N)/owner`), keyed
    // by the same Id groupId() hands out. Overridable per test so a declared
    // owner_group naming a CUSTOM group, rather than a built-in, can be
    // modelled as already correct -- exercising resolveGroupOwner without
    // also exercising CSOM ProcessQuery correction, which this mock does not
    // apply. Default (Id 9, i.e. every unconfigured name) matches the
    // built-in Site Owners every other test in this file declares as
    // owner_group, so their existing mismatch-free behaviour is unchanged.
    const GROUP_CURRENT_OWNER = { 9: { Id: 3, Title: 'Site Owners', PrincipalType: 8 } };
    const currentOwnerFor = (id) => (
      GROUP_CURRENT_OWNER[id] || { Id: 3, Title: 'Site Owners', PrincipalType: 8 }
    );
    // Role definitions (custom permission levels). Task 3: the single line
    // this replaces answered every roledefinitions read alike, whether it
    // was the existence probe, a getbyname resolve, or a by-Id read-back --
    // so a MERGE could never be observed landing or failing to land.
    // Seeded with the one level the fixture declares, 'Schema Manager'
    // (test/fixtures/sharepoint-mapping.yaml:72). Its default Description
    // already carries THIS family's marker, matching every other pre-existing
    // object in this harness's "adopted site" fiction: a level a prior run of
    // the same family already created and stamped, so a fresh run reconciles
    // it rather than refusing it. ROLE_DEF_ABSENT and
    // ROLE_DEF_DESCRIPTION_OVERRIDE let #224's adoption-gate tests put the
    // level through the create path, or give it an unmarked or
    // other-family-marked Description, without touching every other test.
    // ROLE_DEF_DROP_FIELD_ON_WRITE follows GROUP_DROP_FIELD_ON_WRITE: it
    // names one field the MERGE accepts but does not store, so a later test
    // can prove a permission-level read-back fails closed.
    let nextRoleDefId = 2;
    const ROLE_DEF_DROP_FIELD_ON_WRITE = null;
    const ROLE_DEF_SETTINGS_KEYS = ['Description', 'High', 'Low'];
    const ROLE_DEF_ABSENT = false;
    const ROLE_DEF_DESCRIPTION_OVERRIDE = null;
    const ROLE_DEF_STATE = ROLE_DEF_ABSENT ? {} : {
      'Schema Manager': {
        Id: 1,
        Description: ROLE_DEF_DESCRIPTION_OVERRIDE == null
          ? 'Test permission level. '
            + 'Provisioned by dbml-sharepoint from simple-test for level Schema Manager.'
          : ROLE_DEF_DESCRIPTION_OVERRIDE,
        BasePermissions: { High: '0', Low: '2049' },
      },
    };
    const roleDefState = (name) => (ROLE_DEF_STATE[name] ||= {
      Id: nextRoleDefId++, Description: '', BasePermissions: { High: '0', Low: '0' },
    });
    // Decode idiom shared with listOf/groupNameOf: odataName DOUBLES an
    // apostrophe and encodeURIComponent leaves it alone, so undo percent
    // encoding first and then the doubling.
    const roleDefFilterNameOf = (url) => {
      const raw = (url.match(/\$filter=Name eq '(.*?)'/) || [])[1];
      return raw == null ? raw : decodeURIComponent(raw).replace(/''/g, "'");
    };
    const roleDefByNameOf = (url) => {
      const raw = (url.match(/roledefinitions\/getbyname\('(.*?)'\)/) || [])[1];
      return raw == null ? raw : decodeURIComponent(raw).replace(/''/g, "'");
    };
    // The web/lists/getbytitle('X')/roleassignments enumeration
    // _acls.js.j2 reads to decide what a list already has bound, keyed by
    // list title. Configurable per title so a later test can exercise the
    // __next pagination path; empty by default, matching every other
    // probe's brand-new-site fiction ('nothing bound yet') so an
    // unconfigured run only ever adds.
    const ROLE_ASSIGNMENT_PAGES = {};
    const roleAssignmentPages = (listTitle) => ROLE_ASSIGNMENT_PAGES[listTitle] || [[]];
    // Per-list Title state, mutated by MERGEs exactly as SharePoint would.
    const titles = {};
    const titleState = (listTitle) => (titles[listTitle] ||= {
      Sealed: true, Required: true, Description: '', DefaultValue: null,
    });
    const titleField = (listTitle) => ({
      Id: '11111111-1111-1111-1111-111111111111',
      InternalName: 'Title', Title: 'Title', TypeAsString: 'Text',
      EnforceUniqueValues: false, Indexed: false, ReadOnlyField: false,
      CustomFormatter: null, ...titleState(listTitle),
    });
    // Created fields persist, so the run converges instead of failing
    // "missing after creation" and aborting before PROTECTION. Without
    // this the mock could never execute a phase past list creation.
    const TYPE_BY_KIND = { 2: 'Text', 3: 'Note', 4: 'DateTime', 6: 'Choice',
      7: 'Lookup', 8: 'Boolean', 9: 'Number', 11: 'URL', 20: 'User',
      17: 'Calculated' };
    const created = {};   // `${list} ${title}` -> shape
    // The list title out of a URL, back in the spelling the declaration uses.
    // `[^']+` would stop at the first apostrophe of an OData-escaped title
    // (odataName doubles `'`, and encodeURIComponent does not touch it), so
    // `O'Brien Register` keyed as `O` and every per-list mock state silently
    // went to the wrong bucket. Non-greedy to the first `')`, then undo the
    // two encodings in the order odataName applied them: percent first,
    // doubling second.
    const listOf = (url) => {
      const raw = (url.match(/getbytitle\('(.*?)'\)/) || [])[1];
      return raw == null ? raw : decodeURIComponent(raw).replace(/''/g, "'");
    };
    const views = {};
    const viewOf = (url) => {
      const match = url.match(/\/views\/getbytitle\('([^']+)'\)/);
      return match && match[1];
    };
    const viewState = (listTitle, title = 'All Items') => (
      views[`${listTitle} ${title}`] ||= {
        Id: '44444444-4444-4444-4444-444444444444',
        Title: title, DefaultView: true, RowLimit: 30, ViewQuery: '',
        Hidden: false, PersonalView: false, CustomFormatter: null,
        ServerRelativeUrl: `/sites/test/Lists/${listTitle}/AllItems.aspx`,
        ViewFields: { Items: { results: ['Title'] } },
      }
    );
    const fieldShape = (listTitle, name, b) => ({
      Id: '33333333-3333-3333-3333-333333333333',
      InternalName: name, Title: name,
      TypeAsString: TYPE_BY_KIND[b.FieldTypeKind] || 'Text',
      Description: b.Description == null ? '' : b.Description,
      Required: b.Required === true,
      EnforceUniqueValues: b.EnforceUniqueValues === true,
      Indexed: b.EnforceUniqueValues === true,
      ReadOnlyField: b.FieldTypeKind === 17,
      Sealed: false,
      DefaultValue: b.DefaultValue == null ? null : b.DefaultValue,
      CustomFormatter: b.CustomFormatter == null ? null : b.CustomFormatter,
      __body: b,
    });
    const body = (url, opts) => {
      if (url.includes('contextinfo')) {
        return { d: { GetContextWebInformation: {
          FormDigestValue: 'digest', FormDigestTimeoutSeconds: 1800 } } };
      }
      if (url.toLowerCase().includes('effectivebasepermissions')) {
        return { d: { EffectiveBasePermissions: { High: 4294967295, Low: 4294967295 } } };
      }
      // A field probe or enumeration. Title exists from the start (it is
      // the adopted, sealed one); everything else appears once created.
      // Checked BEFORE the list probe, whose own $select also names
      // ValidationFormula.
      if (url.includes('/fields')) {
        const listTitle = listOf(url);
        if (url.includes('ClientValidationFormula')) {
          const probed = (url.match(/getbyinternalnameortitle\('([^']+)'\)/) || [])[1];
          const f = created[`${listTitle} ${probed}`] || {};
          return { d: {
            ClientValidationFormula: f.__cvf == null ? null : f.__cvf,
            ClientValidationMessage: null,
            ValidationFormula: f.__vf == null ? null : f.__vf,
            ValidationMessage: f.__vm == null ? null : f.__vm } };
        }
        const named = (url.match(/getbyinternalnameortitle\('([^']+)'\)/) || [])[1];
        if (named === 'Title') return { d: titleField(listTitle) };
        if (named) {
          const f = created[`${listTitle} ${named}`];
          if (!f) return { error: { code: '-2147024809, System.ArgumentException' } };
          // A derived-property probe (MaxLength, Choices, DisplayFormat...)
          // names none of the shape columns; echo what the field was
          // created with, which is what the declaration asked for.
          if (!url.includes('InternalName')) return { d: f.__body };
          return { d: f };
        }
        const own = Object.entries(created)
          .filter(([k]) => k.startsWith(`${listTitle} `))
          .map(([, v]) => v);
        return { d: { results: [titleField(listTitle), ...own] } };
      }
      // Principals: enough shape to get PREPARE past 1.2/1.3 and reach the
      // maintenance unseal at 1.4. Before this, the runtime test had never
      // executed a phase beyond the read-only preflight.
      if (url.includes('AssociatedOwnerGroup') || url.includes('AssociatedMemberGroup')
          || url.includes('AssociatedVisitorGroup')) {
        return { d: { Id: 3, Title: 'Site Owners', PrincipalType: 8 } };
      }
      // A governed group's current owner (`web/sitegroups(N)/owner`), keyed
      // by the Id in the URL so two groups in the same run can carry
      // different current owners. See GROUP_CURRENT_OWNER above.
      if (url.includes('/owner')) {
        const idMatch = url.match(/sitegroups\((\d+)\)\/owner/);
        return { d: currentOwnerFor(idMatch ? Number(idMatch[1]) : 9) };
      }
      // The site-group enumeration the security phase uses to decide
      // create-vs-adopt without a per-group 404 probe. Checked before the
      // by-name probe below, whose URL also contains 'sitegroups' but not
      // this exact query shape.
      if (url.includes('web/sitegroups?')) {
        return { d: { results: KNOWN_GROUP_NAMES.map((name) => ({ Title: name })) } };
      }
      // A group's own membership, by NAME (the shape countGroupMembers and
      // require_empty_at_deploy both read) and paginated the same way the
      // reader-enrolment mock pages sitegroups(N)/users: page 0 unless the
      // caller followed a __next this mock handed out.
      if (url.includes('/users')) {
        const name = groupNameOf(url);
        // INFERRED, NOT MEASURED: the parent by-name GET answers 404 for an
        // absent group (measured; see surveyGroup in
        // _security_principals.js.j2), but what
        // this /users sub-resource answers for an absent group has not been
        // probed. 404 is used because the parent does and because either
        // status fails the template closed; a future probe should confirm
        // or correct this.
        if (!groupIsKnown(name)) {
          return { error: { code: '-2147024809, System.ArgumentException', status: 404 } };
        }
        const pages = groupMemberPages(name);
        const marked = /[?&]page=(\d+)/.exec(url);
        const page = marked ? Number(marked[1]) : 0;
        const payload = { d: { results: pages[page] || [] } };
        if (page + 1 < pages.length) {
          payload.d.__next =
            `https://example.sharepoint.com/_api/web/sitegroups/getbyname('${encodeURIComponent(name)}')`
            + `/users?$select=Id&$top=5000&page=${page + 1}`;
        }
        return payload;
      }
      if (url.includes('sitegroups/getbyname')) {
        const name = groupNameOf(url);
        // MEASURED (surveyGroup in _security_principals.js.j2): a by-name GET
        // for a site group that is not there answers 404. The role-definition
        // getbyname is NOT the same: that one answers 500 for an absent level,
        // which is why the template probes levels by $filter instead.
        // groupIsKnown is checked
        // before groupState(name), whose `||=` would otherwise auto-vivify
        // an absent name into "existing" the instant it is read.
        if (!groupIsKnown(name)) {
          return { error: { code: '-2147024809, System.ArgumentException', status: 404 } };
        }
        return { d: { Id: groupId(name), Title: name, PrincipalType: 8, ...groupState(name) } };
      }
      // Every roledefinitions read shares that substring, so the most
      // specific shape is checked first: the $filter existence probe (by
      // Name), a by-Id read-back, getbyname (both the MERGE target below
      // and what resolveRoleDefId GETs directly), then the bare collection
      // endpoint, which only a create POST reaches -- the write-application
      // block below has already recorded the new state by the time this
      // runs, so it is echoed back the way SharePoint would.
      if (url.includes('roledefinitions')) {
        const notFound = { error: { code: '-2147024809, System.ArgumentException' } };
        if (url.includes('$filter=Name')) {
          const state = ROLE_DEF_STATE[roleDefFilterNameOf(url)];
          const row = state ? [{ Id: state.Id, Description: state.Description }] : [];
          return { d: { results: row } };
        }
        const byId = url.match(/roledefinitions\((\d+)\)/);
        if (byId) {
          const state = Object.values(ROLE_DEF_STATE).find((s) => String(s.Id) === byId[1]);
          return state ? { d: state } : notFound;
        }
        if (url.includes('getbyname')) {
          const state = ROLE_DEF_STATE[roleDefByNameOf(url)];
          return state ? { d: state } : notFound;
        }
        if (opts && opts.body) {
          const parsed = JSON.parse(opts.body);
          const state = ROLE_DEF_STATE[parsed.Name];
          if (state) return { d: state };
        }
        return { d: { results: [] } };
      }
      // A list's role-assignment enumeration (Member + RoleDefinitionBindings),
      // paginated the same way the group membership mock pages
      // sitegroups/.../users: page 0 unless the caller followed a __next
      // this mock handed out.
      if (url.includes('/roleassignments')) {
        const listTitle = listOf(url);
        const pages = roleAssignmentPages(listTitle);
        const marked = /[?&]page=(\d+)/.exec(url);
        const page = marked ? Number(marked[1]) : 0;
        const payload = { d: { results: pages[page] || [] } };
        if (page + 1 < pages.length) {
          payload.d.__next =
            `https://example.sharepoint.com/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')`
            + `/roleassignments?$expand=Member,RoleDefinitionBindings&page=${page + 1}`;
        }
        return payload;
      }
      // The adopted list starts with SharePoint's built-in Title-only All
      // Items view. View writes below mutate this state so exact field/query
      // readback exercises the generated recovery-view behavior.
      if (url.includes('/views?')) {
        const listTitle = listOf(url);
        return { d: { results: [viewState(listTitle)] } };
      }
      if (url.includes('/views/getbytitle')) {
        const state = viewState(listOf(url), viewOf(url));
        if (url.includes('/viewfields')) return { d: state.ViewFields };
        return { d: state };
      }
      // The single list ENUMERATION. This mock's fiction is "any list probe
      // succeeds", which an enumeration cannot express, since it would have to
      // know the declared names. Refusing it exercises the documented
      // fallback in ensureKnownListTitles: enumeration unavailable, probe
      // per list. The fast path itself is NOT covered here.
      if (url.includes('web/lists?')) return { error: { code: 'enumeration-not-mocked' } };
      // A list probe: the list exists, matching the declared shape.
      if (url.includes('getbytitle') && url.includes('BaseTemplate')) {
        return { d: {
          Id: '22222222-2222-2222-2222-222222222222',
          Title: 'adopted', BaseTemplate: 100, ContentTypesEnabled: false,
          Description: listDescription(listOf(url)),
          EnableVersioning: true, EnableMinorVersions: false,
          MajorVersionLimit: 500, ValidationFormula: null, ValidationMessage: null } };
      }
      return { d: { results: [] } };
    };
    globalThis.fetch = async (url, opts = {}) => {
      const u = String(url);
      // body is null, never absent: JSON.stringify drops an undefined key,
      // and the Python side reads c['body'] unconditionally.
      calls.push({ url: u, method: opts.method || 'GET',
                   body: opts.body === undefined ? null : opts.body });
      // Apply writes, exactly as SharePoint would, so readbacks converge.
      if ((opts.method || 'GET') === 'POST' && opts.body && u.includes('/fields')) {
        const parsed = JSON.parse(opts.body);
        const listTitle = listOf(u);
        const named = (u.match(/getbyinternalnameortitle\('([^']+)'\)/) || [])[1];
        if (named === 'Title') {
          for (const k of ['Sealed', 'Required', 'Description', 'DefaultValue']) {
            if (parsed[k] !== undefined) titleState(listTitle)[k] = parsed[k];
          }
        } else if (named) {
          const key = `${listTitle} ${named}`;
          const f = created[key];
          if (f) {
            if (parsed.Sealed != null) f.Sealed = parsed.Sealed;
            if (parsed.ClientValidationFormula != null) f.__cvf = parsed.ClientValidationFormula;
            if (parsed.ValidationFormula != null) f.__vf = parsed.ValidationFormula;
            if (parsed.ValidationMessage != null) f.__vm = parsed.ValidationMessage;
            for (const k of ['Description', 'Required', 'DefaultValue', 'CustomFormatter']) {
              if (parsed[k] !== undefined) f[k] = parsed[k];
            }
          }
        } else if (parsed.Title) {
          created[`${listTitle} ${parsed.Title}`] = fieldShape(listTitle, parsed.Title, parsed);
        }
      }
      // A MERGE onto the LIST object itself. The URL ends at getbytitle(...)
      // with nothing after it, which is what separates a list write from a
      // field or view write under the same list. `[^/]*` rather than `[^']+`
      // for the title: odataName DOUBLES an apostrophe, so `[^']+` would miss
      // every list whose name has one and this mock would silently drop the
      // write while answering 200; see _LIST_WRITE_URL for the full note.
      // IGNORE_DESCRIPTION_WRITES drops it on purpose instead: a write
      // SharePoint reports as 200 and discards, which is the only state in
      // which the read-back can be watched failing.
      if ((opts.method || 'GET') === 'POST' && opts.body
          && /getbytitle\('[^/]*'\)$/.test(u)) {
        const parsed = JSON.parse(opts.body);
        if (parsed.Description !== undefined && !IGNORE_DESCRIPTION_WRITES) {
          LIST_DESCRIPTIONS[listOf(u)] = parsed.Description;
        }
      }
      if ((opts.method || 'GET') === 'POST' && u.includes('/views/getbytitle')) {
        const state = viewState(listOf(u), viewOf(u));
        if (u.includes('/viewfields/removeallviewfields')) {
          state.ViewFields.Items.results = [];
        } else {
          const added = (u.match(/addviewfield\('([^']+)'\)/) || [])[1];
          if (added) {
            state.ViewFields.Items.results.push(added);
          } else if (opts.body) {
            const parsed = JSON.parse(opts.body);
            for (const key of ['Title', 'DefaultView', 'Hidden', 'RowLimit', 'ViewQuery']) {
              if (parsed[key] !== undefined) state[key] = parsed[key];
            }
          }
        }
      }
      // A group create (POST to .../web/sitegroups) or a MERGE onto the
      // group object itself (POST to .../sitegroups/getbyname('...') with
      // nothing after the closing paren, so a membership write to the same
      // group's /users does not match). Mutates GROUP_STATE so
      // verifyGroupSettings's read-back sees what this write actually
      // stored, applying GROUP_DROP_FIELD_ON_WRITE and
      // GROUP_COERCE_AUTO_ACCEPT.
      if ((opts.method || 'GET') === 'POST' && opts.body
          && (u.endsWith('/sitegroups') || /sitegroups\/getbyname\('.*'\)$/.test(u))) {
        const parsed = JSON.parse(opts.body);
        if (parsed.__metadata && parsed.__metadata.type === 'SP.Group') {
          const name = u.endsWith('/sitegroups') ? parsed.Title : groupNameOf(u);
          const state = groupState(name);
          for (const key of GROUP_SETTINGS_KEYS) {
            if (parsed[key] === undefined || key === GROUP_DROP_FIELD_ON_WRITE) continue;
            state[key] = parsed[key];
          }
          if (GROUP_COERCE_AUTO_ACCEPT && !state.AllowRequestToJoinLeave) {
            state.AutoAcceptRequestToJoinLeave = false;
          }
        }
      }
      // A permission-level create (POST to .../web/roledefinitions) or a
      // MERGE onto the definition itself (POST to
      // .../roledefinitions/getbyname('...') with nothing after the closing
      // paren, matching the group write above). Mutates ROLE_DEF_STATE so a
      // later by-Id or by-name read-back sees what this write actually
      // stored, applying ROLE_DEF_DROP_FIELD_ON_WRITE the same way
      // GROUP_DROP_FIELD_ON_WRITE models a write the tenant 200s and
      // discards.
      if ((opts.method || 'GET') === 'POST' && opts.body
          && (u.endsWith('/roledefinitions') || /roledefinitions\/getbyname\('.*'\)$/.test(u))) {
        const parsed = JSON.parse(opts.body);
        if (parsed.__metadata && parsed.__metadata.type === 'SP.RoleDefinition') {
          const name = u.endsWith('/roledefinitions') ? parsed.Name : roleDefByNameOf(u);
          const state = roleDefState(name);
          const sent = {
            Description: parsed.Description,
            High: parsed.BasePermissions && parsed.BasePermissions.High,
            Low: parsed.BasePermissions && parsed.BasePermissions.Low,
          };
          for (const key of ROLE_DEF_SETTINGS_KEYS) {
            if (sent[key] === undefined || key === ROLE_DEF_DROP_FIELD_ON_WRITE) continue;
            if (key === 'Description') state.Description = sent[key];
            else state.BasePermissions[key] = sent[key];
          }
        }
      }
      const payload = body(u, opts);
      const absent = payload && payload.error;
      // Most absence mocks in this file don't carry a measured status and
      // default to 400. The site-group absence mocks above set one
      // explicitly (404), matching what is measured (or inferred) for them.
      const status = absent ? (payload.error.status || 400) : 200;
      return {
        ok: !absent, status,
        headers: { get: () => null },
        json: async () => payload,
        text: async () => JSON.stringify(payload),
      };
    };
    globalThis.__calls = calls;
""")


def _run_deploy(harness: str, tail: str) -> str:
    script = harness + "\n" + _deploy_js().replace("})();", tail).replace(
        "(async () => {", "((async () => {", 1,
    )
    return _run(script)


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_sealed_builtin_title_does_not_abort_every_list() -> None:
    """`assertFieldImmutableShape` throws when a field is sealed and
    `field.seal` is falsy. Both synthetic Title objects omitted the key, so
    against a site whose Title is sealed EVERY list failed preflight, and
    the tool could not self-heal, because the maintenance unseal walks
    declared columns only and Title is not one. A site that ever sealed
    Title was permanently un-deployable."""
    output = _run_deploy(
        _ADOPTED_HARNESS,
        "}))().then(r => console.log('__RESULT__' + JSON.stringify(r)))",
    )
    line = next(
        (ln for ln in output.splitlines() if ln.startswith("__RESULT__")), None,
    )
    assert line is not None, f"deploy.js did not return a summary:\n{output[-3000:]}"
    summary = json.loads(line.removeprefix("__RESULT__"))
    seal_errors = [
        err for err in (summary.get("errors") or []) if "sealed" in str(err)
    ]
    assert not seal_errors, f"a sealed built-in Title aborted the run: {seal_errors}"


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_sealed_title_is_unsealed_for_the_run() -> None:
    """Not aborting is not enough to make the site deployable: Phase 2.1
    writes list.title_patch to Title, and a sealed column discards writes.
    The maintenance unseal walked declared columns only, and Title is not
    one, so the run could never converge. It must open Title too."""
    output = _run_deploy(
        _ADOPTED_HARNESS,
        "}))().then(() => console.log('__CALLS__' + JSON.stringify(globalThis.__calls)))",
    )
    line = next(
        (ln for ln in output.splitlines() if ln.startswith("__CALLS__")), None,
    )
    assert line is not None, f"harness produced no call log:\n{output[-3000:]}"
    calls = json.loads(line.removeprefix("__CALLS__"))
    seal_writes = [
        json.loads(c["body"])["Sealed"]
        for c in calls
        if c["method"] == "POST" and c.get("body")
        and "getbyinternalnameortitle('Title')" in c["url"]
        and "Sealed" in c["body"]
    ]
    assert False in seal_writes, "a sealed Title was never unsealed for the run"


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_the_adopted_run_reaches_the_write_phases() -> None:
    """Guards the reach of the harness itself.

    The original mock answered every field probe as absent and every list
    probe as malformed, so the run aborted in the read-only preflight: no
    phase past 1.1 had ever executed in a test, which is how a bug in the
    Phase 2.1 field reconcile shipped in a green suite. If a future change
    quietly shortens this run, the coverage disappears silently, so the
    reach is asserted rather than assumed."""
    output = _run_deploy(
        _ADOPTED_HARNESS,
        "}))().then(r => console.log('__RESULT__' + JSON.stringify(r)))",
    )
    reached = [ln.split("Starting Phase ")[1][:3] for ln in output.splitlines()
               if "Starting Phase " in ln]
    # `unseal` by key, not by number: the enterprise-reader step renumbered
    # it once already, and this test is about REACH, not about numbering
    # (which test_phases pins).
    for phase in ("1.1", "1.2", "1.3", pn("unseal"), "2.1"):
        assert phase in reached, f"phase {phase} not reached: {reached}"


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_protection_restores_only_the_titles_prepare_unsealed(tmp_path: Path) -> None:
    """The tool does not own Title's seal state, so a run that unseals one
    must hand back what it found: it must neither seal a Title it found
    unsealed nor leave open one it opened to write."""
    js = _declared_deploy_js(
        tmp_path,
        """
        form_visibility:
          Escalation:
            columns:
              Note: hidden
        """,
    )
    script = _ADOPTED_HARNESS + "\n" + js.replace(
        "})();", "}))().then(() => console.log('__CALLS__' + JSON.stringify(globalThis.__calls)))",
    ).replace("(async () => {", "((async () => {", 1)
    line = next(
        (ln for ln in _run(script).splitlines() if ln.startswith("__CALLS__")), None,
    )
    assert line is not None, "harness produced no call log"
    calls = json.loads(line.removeprefix("__CALLS__"))
    seal_writes = [
        json.loads(c["body"])["Sealed"]
        for c in calls
        if c["method"] == "POST" and c.get("body")
        and "getbyinternalnameortitle('Title')" in c["url"]
        and "Sealed" in c["body"]
    ]
    assert seal_writes[0] is False, f"PREPARE did not unseal Title: {seal_writes}"
    assert seal_writes[-1] is True, f"the run left Title unsealed: {seal_writes}"


# A POST to the LIST object itself: the path ends at getbytitle(...) with
# nothing after it. Anchored deliberately. A FIELD MERGE is a POST to
# `web/lists/getbytitle('X')/fields/getbyinternalnameortitle('Y')` and
# routinely carries a Description of its own (every column with a note has
# one), so a filter that only asks for `web/lists` in the URL counts column
# descriptions as list writes, and then no run can ever be observed NOT
# writing a list description, which is half of what these tests measure.
#
# The title is matched as `[^/]*`, NOT `[^']+` and NOT `.*`, and both of the
# rejected spellings are wrong in a way that passes:
#
#   [^']+  cannot match an OData-escaped apostrophe. `odataName`
#          (`_site_guard.js.j2`) DOUBLES `'` and encodeURIComponent leaves it
#          alone, so a list called `O'Brien Register` arrives as
#          getbytitle('O''Brien%20Register'), no match, so the idempotence
#          test observes zero writes for the happiest of reasons and passes.
#   .*     matches too much: greedy backtracking lets it swallow
#          `X')/fields/getbyinternalnameortitle('Y` and call a FIELD write a
#          list write, which is the false positive this anchor exists to stop.
#
# A SharePoint list title cannot contain `/`, and encodeURIComponent would
# percent-encode one anyway, so "no slash after the opening quote" separates
# the list object from everything nested under it. Both directions are pinned
# by test_the_list_write_matcher_survives_an_apostrophe.
_LIST_WRITE_URL = re.compile(r"web/lists/getbytitle\('[^/]*'\)$")


def _description_writes(calls: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Every request that MERGEs a Description onto a list."""
    return [
        c for c in calls
        if c["method"] == "POST" and c["body"] and "Description" in c["body"]
        and _LIST_WRITE_URL.search(c["url"])
    ]


def _declared_list_descriptions(
    tmp_path: Path, prefix: str = DEFAULT_PREFIX,
) -> dict[str, str]:
    """List title -> the Description `_declared_deploy_js` declares for it.

    Read out of the generator, off the SAME pack the script is built from,
    rather than re-spelled here: a second copy of the fixture would drift,
    and a declared-against-live test comparing two different fixtures proves
    nothing. Returns a mapping rather than one string because the marker
    embeds the entity name, so no single value can be "the declared
    description" for more than one list.
    """
    from dbml_sharepoint.generators.jsgen import build_schema_json

    schema, bundle = _declared_pack(tmp_path, "", prefix)
    schema_json = build_schema_json(schema, bundle, "default")
    return {entry["title"]: entry["description"] for entry in schema_json["lists"]}


def _run_adopted_deploy(
    tmp_path: Path,
    list_description: str | dict[str, str],
    *,
    ignore_description_writes: bool = False,
    prefix: str = DEFAULT_PREFIX,
) -> tuple[dict[str, Any], list[dict[str, Any]], str]:
    """Run the emitted deploy against a site whose lists already exist.

    Built on `_declared_deploy_js`, not on the shipped `simple.dbml` fixture.
    That matters for the abort assertion: `test_a_declared_run_completes_every
    _phase_cleanly` pins this schema as finishing with NO errors and NO abort
    against `_ADOPTED_HARNESS`, whereas the simple fixture's adopted run
    already aborts on `phase-1-schema-errors` (the mock is too thin for its
    renamed and indexed columns). On that base `summary['aborted']` is truthy
    no matter what the description does, and the read-back test could not
    fail, which is worse than not having it.

    `list_description` is what the site HOLDS before the run: one string for
    every list, or a per-title mapping. `ignore_description_writes` makes the
    mock accept the MERGE with a 200 and keep serving the old value (a
    silently discarded write, which is the only state in which the read-back
    can be watched failing).

    `prefix` reaches the mapping's list-title prefix, which is how a caller
    deploys to a list whose title needs OData escaping.

    Returns (summary, calls, output). The list phase must actually have
    started: otherwise a "nothing was written" assertion would pass against a
    run that aborted in the preflight and never reached the reconcile at all.
    """
    held = (
        dict.fromkeys(
            _declared_list_descriptions(tmp_path, prefix), list_description,
        )
        if isinstance(list_description, str) else dict(list_description)
    )
    harness = _ADOPTED_HARNESS.replace(
        "const LIST_DESCRIPTIONS = {};",
        f"const LIST_DESCRIPTIONS = {json.dumps(held)};",
    ).replace(
        "const IGNORE_DESCRIPTION_WRITES = false;",
        f"const IGNORE_DESCRIPTION_WRITES = {json.dumps(ignore_description_writes)};",
    )
    script = harness + "\n" + _declared_deploy_js(tmp_path, "", prefix).replace(
        "})();",
        "}))().then(r => { console.log('__RESULT__' + JSON.stringify(r));"
        " console.log('__CALLS__' + JSON.stringify(globalThis.__calls)); })",
    ).replace("(async () => {", "((async () => {", 1)
    output = _run(script)
    result_line = next(
        (ln for ln in output.splitlines() if ln.startswith("__RESULT__")), None,
    )
    assert result_line is not None, f"deploy.js did not return a summary:\n{output[-3000:]}"
    calls_line = next(
        (ln for ln in output.splitlines() if ln.startswith("__CALLS__")), None,
    )
    assert calls_line is not None, f"harness produced no call log:\n{output[-3000:]}"
    # By KEY, not by number: phase numbers derive from position and renumber
    # themselves the moment anybody inserts a phase, and a hardcoded '2.1'
    # would then silently stop guarding reach.
    assert f"Starting Phase {pn('lists')}" in output, (
        f"the list reconcile phase never ran:\n{output[-3000:]}"
    )
    return (
        json.loads(result_line.removeprefix("__RESULT__")),
        json.loads(calls_line.removeprefix("__CALLS__")),
        output,
    )


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_an_existing_list_with_the_wrong_description_is_corrected(
    tmp_path: Path,
) -> None:
    """The adoption path is the one that matters.

    A list provisioned before markers existed, or one an owner edited, holds
    a description discovery cannot match. Creation-only writing leaves it
    that way forever and reports success.
    """
    summary, calls, output = _run_adopted_deploy(
        tmp_path, "something an owner typed",
    )
    writes = _description_writes(calls)
    assert writes, (
        "an existing list kept a description with no marker and the run "
        f"reported success; it is now invisible to fleet reporting\n{output[-2000:]}"
    )
    assert "Provisioned by dbml-sharepoint" in writes[0]["body"]
    # The repair has to CONVERGE, not merely be attempted: a MERGE whose
    # read-back then failed would satisfy the assertions above while leaving
    # the operator with an aborted run.
    assert summary.get("aborted") is None, summary
    assert summary.get("errors") == [], summary["errors"]


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_correct_description_is_not_rewritten(tmp_path: Path) -> None:
    """Idempotence: a re-paste must not churn every list it looks at."""
    declared = _declared_list_descriptions(tmp_path)
    # Without this the test is vacuous: an empty declared description would
    # also never be rewritten, and nothing else here would notice the marker
    # had gone missing from the generator entirely.
    assert declared and all(
        "Provisioned by dbml-sharepoint" in value for value in declared.values()
    ), declared
    summary, calls, output = _run_adopted_deploy(tmp_path, declared)
    # Same guard as the sibling test, and for the same reason: "no description
    # was written" is also true of a run that fell over before it got there.
    # Without this the test would keep passing through an unrelated breakage
    # that stopped the reconcile from executing at all.
    assert summary.get("aborted") is None, summary
    assert summary.get("errors") == [], summary["errors"]
    assert not _description_writes(calls), (
        f"a list already carrying its declared description was rewritten"
        f"\n{output[-2000:]}"
    )


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_description_that_does_not_read_back_aborts(tmp_path: Path) -> None:
    """`AGENTS.md`: anything that writes must read back and verify.

    A MERGE that returns 200 while the stored value stays stale is the exact
    shape this repository exists to catch -- the deploy reports success and
    the list is still undiscoverable.
    """
    summary, calls, output = _run_adopted_deploy(
        tmp_path, "stale", ignore_description_writes=True,
    )
    assert _description_writes(calls), f"nothing was even attempted\n{output[-2000:]}"
    assert summary.get("aborted"), (
        f"the description never took and the run still reported success"
        f"\n{output[-2000:]}"
    )
    assert "did not retain its declared Description" in output, output[-2000:]


# A list title carrying the one character OData escapes by DOUBLING rather
# than by percent-encoding. `prefix` is the knob because the rest of a list
# title is the DBML table name, which the parser will not let hold one.
_APOSTROPHE_PREFIX = "prefix: \"O'Brien \""


def test_the_list_write_matcher_survives_an_apostrophe() -> None:
    """The matcher must be exact in BOTH directions, and neither is obvious.

    `odataName` doubles an apostrophe and encodeURIComponent leaves it alone,
    so a title like `O'Brien Register` reaches the wire as
    getbytitle('O''Brien%20Register'). A `[^']+` title pattern stops at the
    first quote and matches nothing, which does not look like a broken
    matcher, it looks like a run that correctly wrote nothing, and the
    idempotence test passes for that reason forever.

    `.*` is the other trap: greedy backtracking lets it swallow the rest of
    the path, so a FIELD MERGE (which routinely carries its own Description)
    is counted as a list write.

    Asserted over the real URL shapes rather than over prose, because both
    failures are the kind that get reasoned about correctly and coded wrongly.
    """
    escaped = "/sites/x/_api/web/lists/getbytitle('O''Brien%20Register')"
    plain = "/sites/x/_api/web/lists/getbytitle('APP_Plain')"
    assert _LIST_WRITE_URL.search(escaped), "an escaped apostrophe was not matched"
    assert _LIST_WRITE_URL.search(plain)
    for nested in (
        f"{escaped}/fields/getbyinternalnameortitle('Note')",
        f"{escaped}/views/getbytitle('All%20Items')",
        f"{plain}/fields/getbyinternalnameortitle('Note')",
    ):
        assert not _LIST_WRITE_URL.search(nested), (
            f"a write nested under the list was counted as a list write: {nested}"
        )


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_description_is_reconciled_on_a_list_whose_title_needs_escaping(
    tmp_path: Path,
) -> None:
    """End-to-end companion to the matcher test above.

    Pins the whole chain against an OData-escaped title at once: the emitted
    script builds the URL, the mock recognises it as a list write and applies
    it, the read-back sees it, and the harness's own `listOf` keys the state
    by the right list. Any one of those regressing to a `[^']+` title pattern
    turns this red, where reasoning about it would just leave the other
    tests quietly passing on a name no fixture happens to use.
    """
    declared = _declared_list_descriptions(tmp_path, _APOSTROPHE_PREFIX)
    assert any("'" in title for title in declared), declared
    summary, calls, output = _run_adopted_deploy(
        tmp_path, "typed by an owner", prefix=_APOSTROPHE_PREFIX,
    )
    writes = _description_writes(calls)
    assert writes, (
        "no list write was seen for a title carrying an apostrophe; either the "
        f"script or the matcher lost it\n{output[-2000:]}"
    )
    assert "Provisioned by dbml-sharepoint" in writes[0]["body"]
    assert summary.get("aborted") is None, summary
    assert summary.get("errors") == [], summary["errors"]


# The adopted site again, but every field CREATION is refused. STRUCTURE
# then records an error per column and takes its early return, the
# designed abort that skips ACL work on a broken schema. It also skips
# PROTECTION, which is where a Title unsealed at 1.4 used to be handed
# back. Only creation is refused: the 1.4 MERGE that unseals Title is a
# write to an existing field and still succeeds.
_ABORTING_HARNESS = _ADOPTED_HARNESS + textwrap.dedent(r"""
    const _passThrough = globalThis.fetch;
    globalThis.fetch = async (url, opts = {}) => {
      const u = String(url);
      const path = u.split('?')[0];
      const creating = (opts.method || 'GET') === 'POST' && opts.body
        && (path.endsWith('/fields') || path.endsWith('/fields/addfield'));
      if (!creating) return _passThrough(url, opts);
      calls.push({ url: u, method: 'POST', body: opts.body });
      const payload = { error: { message: { value: 'field creation refused' } } };
      return {
        ok: false, status: 400,
        headers: { get: () => null },
        json: async () => payload,
        text: async () => JSON.stringify(payload),
      };
    };
""")


_ABORTING_SEALED_FIELD_HARNESS = _ADOPTED_HARNESS + textwrap.dedent(r"""
    // Adopt one ordinary declared field in the sealed state PREPARE finds
    // on a maintained site. Its stale description forces Phase 2.1 to write.
    const adoptedNote = fieldShape('APP_Escalation', 'Note', {
      FieldTypeKind: 2, Required: false, Description: 'stale description', MaxLength: 255,
    });
    adoptedNote.Sealed = true;
    created['APP_Escalation Note'] = adoptedNote;

    const _passThrough = globalThis.fetch;
    globalThis.fetch = async (url, opts = {}) => {
      const u = String(url);
      const parsed = opts.body ? JSON.parse(opts.body) : {};
      const refusingReconcile = (opts.method || 'GET') === 'POST'
        && u.includes("getbyinternalnameortitle('Note')")
        && parsed.Description !== undefined && parsed.Sealed === undefined;
      if (!refusingReconcile) return _passThrough(url, opts);
      calls.push({ url: u, method: 'POST', body: opts.body });
      const payload = { error: { message: { value: 'field reconcile refused' } } };
      return {
        ok: false, status: 400,
        headers: { get: () => null },
        json: async () => payload,
        text: async () => JSON.stringify(payload),
      };
    };
""")


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_run_that_aborts_after_unsealing_a_title_reseals_it() -> None:
    """A failed run must not leave the site less protected than it found it.

    PREPARE unseals an already-sealed built-in Title so the write phases
    can patch it, and PROTECTION hands it back. Every abort between the
    two (schema errors, lookup errors, enrolment errors) returns before
    PROTECTION, so the run ended with a column someone had deliberately
    sealed left open. Restoration must therefore be on the exit path, not
    on the success path."""
    output = _run_deploy(
        _ABORTING_HARNESS,
        "}))().then(r => { console.log('__RESULT__' + JSON.stringify(r));"
        " console.log('__CALLS__' + JSON.stringify(globalThis.__calls)); })",
    )
    result_line = next(
        (ln for ln in output.splitlines() if ln.startswith("__RESULT__")), None,
    )
    assert result_line is not None, f"deploy.js did not return a summary:\n{output[-3000:]}"
    summary = json.loads(result_line.removeprefix("__RESULT__"))
    # Without this the test could pass by never aborting at all, and a
    # run that reaches PROTECTION re-seals for the ordinary reason.
    assert summary.get("aborted"), (
        f"the run did not abort, so it never tested the abort path: {summary}"
    )
    reached = [ln.split("Starting Phase ")[1][:3] for ln in output.splitlines()
               if "Starting Phase " in ln]
    assert pn("unseal") in reached, f"the maintenance unseal never ran: {reached}"
    assert "4.1" not in reached, f"the run reached PROTECTION, so it did not abort early: {reached}"

    calls_line = next(
        (ln for ln in output.splitlines() if ln.startswith("__CALLS__")), None,
    )
    assert calls_line is not None, "harness produced no call log"
    calls = json.loads(calls_line.removeprefix("__CALLS__"))
    seal_writes = [
        json.loads(c["body"])["Sealed"]
        for c in calls
        if c["method"] == "POST" and c.get("body")
        and "getbyinternalnameortitle('Title')" in c["url"]
        and "Sealed" in c["body"]
    ]
    assert seal_writes, "PREPARE never unsealed a Title, so there was nothing to restore"
    assert seal_writes[0] is False, f"PREPARE did not unseal Title: {seal_writes}"
    assert seal_writes[-1] is True, f"the aborted run left Title unsealed: {seal_writes}"


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_run_that_aborts_after_unsealing_a_declared_field_reseals_it(tmp_path: Path) -> None:
    """Title is not special on the exit path: PREPARE opens every declared
    sealed field, so a Phase 2.1 abort must hand every one of them back."""
    js = _declared_deploy_js(tmp_path, "seal_columns: true\n")
    script = _ABORTING_SEALED_FIELD_HARNESS + "\n" + js.replace(
        "})();", "}))().then(r => { console.log('__RESULT__' + JSON.stringify(r));"
        " console.log('__CALLS__' + JSON.stringify(globalThis.__calls)); })",
    ).replace("(async () => {", "((async () => {", 1)
    output = _run(script)

    result_line = next(
        (ln for ln in output.splitlines() if ln.startswith("__RESULT__")), None,
    )
    assert result_line is not None, f"deploy.js did not return a summary:\n{output[-3000:]}"
    summary = json.loads(result_line.removeprefix("__RESULT__"))
    assert summary.get("aborted") == "phase-1-schema-errors"

    calls_line = next(
        (ln for ln in output.splitlines() if ln.startswith("__CALLS__")), None,
    )
    assert calls_line is not None, "harness produced no call log"
    calls = json.loads(calls_line.removeprefix("__CALLS__"))
    seal_writes = [
        json.loads(c["body"])["Sealed"]
        for c in calls
        if c["method"] == "POST" and c.get("body")
        and "getbyinternalnameortitle('Note')" in c["url"]
        and "Sealed" in c["body"]
    ]
    assert seal_writes[0] is False, f"PREPARE did not unseal Note: {seal_writes}"
    assert seal_writes[-1] is True, f"the aborted run left Note unsealed: {seal_writes}"


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_declared_run_completes_every_phase_cleanly(tmp_path: Path) -> None:
    """The end-to-end guard, and the one that gives the others their value.

    The original mock aborted in the read-only preflight, so no phase past
    1.1 had ever executed in a test, which is how a bug in the Phase 2.1
    field reconcile shipped in a green suite. This run adopts an existing
    site, unseals, creates, reconciles declared formulas, seals and seeds,
    and must finish with no errors and no abort. If a future change
    shortens it, the coverage disappears silently unless this fails.
    """
    js = _declared_deploy_js(
        tmp_path,
        """
        form_visibility:
          Escalation:
            columns:
              Note: hidden
        """,
    )
    script = _ADOPTED_HARNESS + "\n" + js.replace(
        "})();", "}))().then(r => console.log('__RESULT__' + JSON.stringify(r)))",
    ).replace("(async () => {", "((async () => {", 1)
    output = _run(script)
    line = next(
        (ln for ln in output.splitlines() if ln.startswith("__RESULT__")), None,
    )
    assert line is not None, f"deploy.js did not return a summary:\n{output[-3000:]}"
    summary = json.loads(line.removeprefix("__RESULT__"))
    assert summary.get("aborted") is None, summary
    assert summary.get("errors") == [], summary["errors"]
    reached = [ln.split("Starting Phase ")[1][:3] for ln in output.splitlines()
               if "Starting Phase " in ln]
    for phase in ("1.1", pn("unseal"), "2.1", "3.1", "4.1", "5.1"):
        assert phase in reached, f"phase {phase} not reached: {reached}"


def test_generated_deploy_js_carries_no_control_characters() -> None:
    """deploy.js is pasted into a browser console by hand.

    A stray control character survives templating, the golden file and
    every text-mode diff (git reports the file as binary and shows
    nothing). Writing this fix, a literal NUL reached a template's
    executable code from an editing tool and rode into the generated
    script; the suite was green. Cheap to assert, invisible otherwise.
    """
    js = _deploy_js()
    stray = sorted({
        ch for ch in js
        if ord(ch) < 32 and ch not in "\n\r\t"
    })
    assert not stray, f"control characters in generated deploy.js: {[hex(ord(c)) for c in stray]}"


def _declared_pack(
    tmp_path: Path, section: str, prefix: str = DEFAULT_PREFIX,
) -> tuple[Any, Any]:
    """The (schema, bundle) behind `_declared_deploy_js`.

    Split out so a test can ask what the generator DECLARES for these lists
    without re-spelling the fixture. A second copy of the schema here would
    drift from the one the script is built from, and the tests that compare
    declared-against-live would then be comparing two different fixtures.

    `prefix` is the only knob that puts an arbitrary character into a LIST
    TITLE. The rest of the title is the DBML table name, which the parser
    constrains. It is what lets a test deploy to a list whose title needs
    OData escaping.
    """
    return pack(
        tmp_path,
        dbml=table("Escalation", ID_PK, "Title nvarchar", "Note nvarchar"),
        mapping=blocks(entities("Escalation"), section),
        prefix=prefix,
    )


def _declared_deploy_js(
    tmp_path: Path, section: str, prefix: str = DEFAULT_PREFIX,
) -> str:
    """deploy.js for an all-text schema that actually declares a formula.

    The shipped fixture declares none, so enforceDeclaredFormulas returns
    before doing anything and cannot be exercised through it. All-Text
    columns keep the run clear of the derived-property probes the mock does
    not answer.

    `section` is whatever extra mapping the test needs. It is dedented, so a
    caller may pass a triple-quoted block indented to match its surrounding
    code. `blocks()` rather than `with_tail()` because every caller opens a
    TOP-LEVEL section here. Nothing nests under the entity, so no
    indentation matters and the two agree.
    """
    from dbml_sharepoint.generators.jsgen import generate_deploy_js
    from dbml_sharepoint.model.release import load_release

    schema, bundle = _declared_pack(tmp_path, section, prefix)
    return generate_deploy_js(
        schema=schema,
        bundle=bundle,
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="s.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
    )


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_overwriting_a_declared_formula_logs_the_prior_value(tmp_path: Path) -> None:
    """`before` was read, compared and discarded; on success nothing was
    logged, so a deploy that removed or rewrote an existing formula left no
    record of what had been there. Under `reconcile: exact` an undeclared
    column's formula is cleared outright, exactly the case where the prior
    value is the only thing anyone would want back."""
    harness = _ADOPTED_HARNESS.replace(
        "ClientValidationFormula: f.__cvf == null ? null : f.__cvf,",
        "ClientValidationFormula: f.__cvf == null ? "
        "\"=if([$WasHere] != '', 'true', 'false')\" : f.__cvf,",
    )
    js = _declared_deploy_js(
        tmp_path,
        """
        form_visibility:
          Escalation:
            columns:
              Note: hidden
        """,
    )
    script = harness + "\n" + js.replace(
        "})();", "}))().then(r => console.log('__RESULT__' + JSON.stringify(r)))",
    ).replace("(async () => {", "((async () => {", 1)
    output = _run(script)
    replaced = [ln for ln in output.splitlines() if "declared formulas" in ln]
    assert replaced, f"no prior value logged:\n{output[-2500:]}"
    assert any("WasHere" in ln for ln in replaced), replaced


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_formula_reconcile_fails_when_sharepoint_drops_validation_message(
    tmp_path: Path,
) -> None:
    harness = _ADOPTED_HARNESS.replace(
        "if (parsed.ValidationMessage != null) f.__vm = parsed.ValidationMessage;",
        "// Simulate SharePoint accepting the MERGE but dropping ValidationMessage.",
    )
    js = _declared_deploy_js(
        tmp_path,
        """
        column_validation:
          Escalation:
            columns:
              Note:
                when: [{ field: Note, op: is_not_null }]
                message: A note is required.
        """,
    )
    script = harness + "\n" + js.replace(
        "})();", "}))().then(r => console.log('__RESULT__' + JSON.stringify(r)))",
    ).replace("(async () => {", "((async () => {", 1)
    output = _run(script)
    assert "did not retain ValidationMessage" in output, output[-3000:]


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_formula_reconcile_fails_when_client_message_is_not_cleared(tmp_path: Path) -> None:
    harness = _ADOPTED_HARNESS.replace(
        "ClientValidationMessage: null,",
        "ClientValidationMessage: 'stale guidance',",
    )
    js = _declared_deploy_js(
        tmp_path,
        """
        form_visibility:
          Escalation:
            columns:
              Note: hidden
        """,
    )
    script = harness + "\n" + js.replace(
        "})();", "}))().then(r => console.log('__RESULT__' + JSON.stringify(r)))",
    ).replace("(async () => {", "((async () => {", 1)
    output = _run(script)
    assert "did not retain ClientValidationMessage" in output, output[-3000:]


def test_the_aggregations_comparison_survives_sharepoints_readback_spacing() -> None:
    """SharePoint returns `<FieldRef Name="X" Type="Sum" />` for the
    `...Type="Sum"/>` it was sent, verified against a live tenant on
    2026-07-29 (test/manual/view-aggregations-probe.js).

    Compared raw, a perfectly correct totals view drifts on EVERY redeploy:
    the phase rewrites the property, reads the same difference back, and
    fails closed. And it does so on the second run, never the first, which
    is the kind of bug that ships.

    This executes the SHIPPED normaliser out of the generated script rather
    than a copy of its logic, because a copy would keep passing after the
    real one changed.
    """
    if NODE is None:
        pytest.skip("node is not installed")
    script = _deploy_js()
    decode = re.search(r"^\s*const xmlDecode = .*?;$", script, re.MULTILINE | re.DOTALL)
    normalise = re.search(r"^\s*const normalizeViewQuery = .*?;$", script, re.MULTILINE)
    assert decode and normalise, "could not extract the normaliser from the generated script"

    sent = '<FieldRef Name="Amount" Type="Sum"/>'
    read_back = '<FieldRef Name="Amount" Type="Sum" />'
    program = (
        f"{decode.group(0)}\n{normalise.group(0)}\n"
        f"const a = normalizeViewQuery({json.dumps(sent)});\n"
        f"const b = normalizeViewQuery({json.dumps(read_back)});\n"
        "console.log(JSON.stringify({ equal: a === b, a, b }));"
    )
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "normalise.js"
        path.write_text(program, encoding="utf-8")
        out = subprocess.run(  # noqa: S603
            [NODE, str(path)], capture_output=True, text=True, check=True, timeout=60,
        )
    result = json.loads(out.stdout.strip())
    assert result["equal"], (
        f"the shipped normaliser does not equalise SharePoint's readback spacing: "
        f"sent normalised to {result['a']!r}, readback to {result['b']!r}"
    )


def test_no_aggregations_comparison_is_made_raw() -> None:
    """The write-side and readback-side comparisons are separate call sites
    and either one left raw reintroduces the never-converging redeploy.

    Asserted as the ABSENCE of any raw comparison rather than the presence
    of two known-good ones: naming the variables would break on a rename
    while saying nothing about a third call site somebody adds later.
    """
    script = _deploy_js()
    raw = re.findall(r"(?<!normalizeViewQuery\()\b\w+\.Aggregations\s*[!=]==", script)
    assert not raw, f"Aggregations compared without the normaliser: {raw}"
    # AggregationsStatus is a plain enum ('On'/'Off') and IS compared raw,
    # asserted so the regex above cannot be "fixed" by wrapping it too.
    assert "AggregationsStatus !== 'On'" in script


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_first_deploy_probes_no_absent_group_or_field_by_name() -> None:
    """A clean run must leave a clean console.

    The browser logs a failed request itself, before the script sees the
    response, and nothing in JavaScript can suppress that, so a handled
    404 still paints red and an operator reads it as a failure. The only
    fix is not to make the request: enumerate once, answer absence
    locally. Lists and views already did; site groups and the field probe
    did not, and a live first deploy showed four red lines because of it.

    The harness answers every enumeration as EMPTY, which is the state of
    a brand-new site, so any by-name probe here is one an operator would
    have seen painted red.

    Covers the two surfaces this harness reaches. The third (a list's
    role assignments by principal id) is asserted structurally instead,
    because the mock's principal resolution never returns an Id and so the
    run never reaches the ACL phase's role-assignment calls. A clause for
    it here would pass while testing nothing.
    """
    script = _HARNESS + "\n" + _deploy_js().replace(
        "})();", "}))().then(() => console.log('__CALLS__' + JSON.stringify(globalThis.__calls)))",
        ).replace("(async () => {", "((async () => {", 1)
    line = next(
        (ln for ln in _run(script).splitlines() if ln.startswith("__CALLS__")), None,
    )
    assert line is not None, "harness produced no call log"
    calls = json.loads(line.removeprefix("__CALLS__"))
    gets = [c["url"] for c in calls if c["method"] == "GET"]
    # The ACL phase resolves a group's Id by name AFTER creating it, so on
    # a real run that request succeeds and is not console noise; the mock
    # creates nothing, which is why it is excluded by its $select rather
    # than by being overlooked.
    by_name = [
        u for u in gets
        if ("sitegroups/getbyname" in u and "$select=Id" not in u)
        or ("getbytitle" in u and "/fields?" in u)
    ]
    assert not by_name, (
        "a first deploy probed by name for something it had already "
        f"enumerated as absent; each is a red console line: {by_name[:5]}"
    )


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_the_unsupported_formula_tolerance_is_scoped_to_clears() -> None:
    """The clearing-only guard is a deliberate hole in fail-closed, so its
    edges are asserted rather than described.

    Extracts the shipped predicate and runs it over every combination: a
    clear (empty or UNMANAGED on both) may be tolerated; anything that SETS
    either formula must not be, or a rejected write to a field type that
    silently drops formulas would be reported as success.
    """
    script = _deploy_js()
    match = re.search(
        r"const clearingOnly = (.*?);\n", script, re.DOTALL,
    )
    assert match, "could not extract the clearingOnly predicate"
    cases = [
        ("", "", True),
        ("__dbmlsp_unmanaged__", "__dbmlsp_unmanaged__", True),
        ("", "__dbmlsp_unmanaged__", True),
        ("=[X]>1", "", False),
        ("", "=[X]>1", False),
        ("=[X]>1", "=[Y]>1", False),
        ("__dbmlsp_unmanaged__", "=[Y]>1", False),
    ]
    program = (
        "const UNMANAGED = '__dbmlsp_unmanaged__';\n"
        "const out = [];\n"
        f"for (const [v, c] of {json.dumps([[a, b] for a, b, _ in cases])}) {{\n"
        "  const field = { validation_formula: v, client_validation_formula: c };\n"
        f"  const clearingOnly = {match.group(1)};\n"
        "  out.push(clearingOnly);\n"
        "}\n"
        "console.log(JSON.stringify(out));"
    )
    assert NODE is not None
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "guard.js"
        path.write_text(program, encoding="utf-8")
        proc = subprocess.run(  # noqa: S603
            [NODE, str(path)], capture_output=True, text=True, check=True, timeout=60,
        )
    got = json.loads(proc.stdout.strip())
    expected = [tolerated for _, _, tolerated in cases]
    assert got == expected, (
        f"clearingOnly must tolerate only clears; got {got}, expected {expected} "
        f"for {[(a, b) for a, b, _ in cases]}"
    )


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_refused_clear_still_retries_the_client_properties() -> None:
    """A MERGE is atomic: the refusal applies none of the body, including a
    ClientValidationFormula clear the same request carried and which a URL
    field does accept. Tolerating the refusal without retrying those would
    report success while a stale show/hide rule stayed live."""
    script = _deploy_js()
    assert "client-only retry also failed" in script, (
        "the tolerant branch must retry the properties the field type accepts"
    )
    # And it must not simply return: the read-back below is what proves the
    # client clear landed.
    tolerant = script[script.index("const clearingOnly ="):]
    tolerant = tolerant[: tolerant.index("const after = await read();")]
    assert "return;" not in tolerant, (
        "the tolerant branch must fall through to the read-back, not return"
    )


# === Enterprise reader enrolment (the reader_enrolment phase) ===
#
# This phase grants Read on a customer's register to a named account, and
# the membership is PERMANENT, unlike the operator's, which the run
# removes on the way out. Two resolutions must never be enrolled: a
# security GROUP (everyone in it gets Read) and one of SharePoint's
# everyone-claims (every user in the tenant gets Read). Neither is visible
# afterwards, because the deploy reads back byte-identical either way.
#
# So every test below RUNS the emitted script and asserts on what the run
# DOES (did it abort, and above all was a membership POST ever issued).
# `assert "PrincipalType !== 1" in js` would pass with the guard sitting in
# a comment; "no POST to sitegroups(N)/users happened" cannot.

_READER_ADDRESS = "svc-reporting@example.org"

# A well-formed resolution of _READER_ADDRESS. Every refusal test below
# starts from this and varies exactly ONE attribute, so the guard under
# test is the only one that can fire. Otherwise deleting it would leave
# the test green for a neighbouring reason and prove nothing.
_RESOLVED_USER: dict[str, Any] = {
    "Id": 42,
    "LoginName": "i:0#.f|membership|svc-reporting@example.org",
    "Title": "Reporting Service",
    "Email": _READER_ADDRESS,
    "PrincipalType": 1,
}

_MEMBERSHIP_URL = re.compile(r"sitegroups\(\d+\)/users")


def _reader_deploy_js(enterprise_reader: str | None = _READER_ADDRESS) -> str:
    """deploy.js for the mapping that declares an enterprise-reader group."""
    from dbml_sharepoint.generators.jsgen import generate_deploy_js
    from dbml_sharepoint.model.mapping_loader import load_mapping
    from dbml_sharepoint.model.parser import parse_dbml
    from dbml_sharepoint.model.release import load_release

    return generate_deploy_js(
        schema=parse_dbml(FIXTURES / "simple.dbml"),
        bundle=load_mapping(FIXTURES / "sharepoint-mapping-with-reader.yaml"),
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="simple.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
        enterprise_reader=enterprise_reader,
    )


def _reader_harness(
    ensure_user: dict[str, Any],
    *,
    members: list[dict[str, Any]] | None = None,
    member_pages: list[list[dict[str, Any]]] | None = None,
    drop_readback: bool = False,
    stray_on_write: dict[str, Any] | None = None,
) -> str:
    """`_ADOPTED_HARNESS` plus the two surfaces the reader phase touches.

    `ensureuser` answers `ensure_user` verbatim. That is the whole point of
    the harness, since what a tenant resolves an address to is exactly what
    the guards have to judge. The flagged group's membership is real state:
    the POST appends to it and the verification re-read sees the result, so
    a run cannot satisfy the read-back by asserting its own success.
    `drop_readback` accepts the POST and drops the write, which is what a
    silently-refused membership looks like from the script's side.

    `stray_on_write` appends a FOREIGN principal at the same moment, which
    is what another administrator adding somebody between the before-read
    and the read-back looks like. It cannot be modelled by seeding
    `members`, because a principal present before the run is caught by the
    gate that runs first -- the whole point is that this one arrives after
    that gate has already passed.

    `member_pages` serves the membership across SEVERAL OData pages, each
    but the last carrying a `__next`. A group whose membership arrives in
    one page cannot distinguish a gate that reads every page from one that
    reads the first and stops, and the second is a gate that a large
    group defeats simply by being large. `members=[...]` is the one-page
    case, and is exactly `member_pages=[[...]]`.
    """
    pages = [list(members or [])] if member_pages is None else [
        list(page) for page in member_pages
    ]
    return _ADOPTED_HARNESS + textwrap.dedent(r"""
        const ENSURED = __ENSURE_USER__;
        const READER_MEMBER_PAGES = __MEMBER_PAGES__;
        const DROP_READBACK = __DROP_READBACK__;
        const STRAY_ON_WRITE = __STRAY_ON_WRITE__;
        const _beforeReader = globalThis.fetch;
        globalThis.fetch = async (url, opts = {}) => {
          const u = String(url);
          const method = opts.method || 'GET';
          const respond = (payload) => {
            calls.push({ url: u, method,
                         body: opts.body === undefined ? null : opts.body });
            return { ok: true, status: 200, headers: { get: () => null },
                     json: async () => payload,
                     text: async () => JSON.stringify(payload) };
          };
          if (u.toLowerCase().includes('/ensureuser')) return respond({ d: ENSURED });
          // Task 6 (security-phase-atomicity): removeReaderEnrollments's
          // drain POSTs here. Checked BEFORE the broader
          // sitegroups(N)/users test below, which this URL also matches --
          // and whose POST branch assumes an add, parsing `opts.body` as
          // JSON. A remove call carries no body, so falling through to that
          // branch throws `JSON.parse(undefined)` instead of modelling the
          // removal.
          const removed = /sitegroups\(\d+\)\/users\/removebyid\((\d+)\)/.exec(u);
          if (removed && method === 'POST') {
            const removedId = Number(removed[1]);
            for (const page of READER_MEMBER_PAGES) {
              const idx = page.findIndex((m) => Number(m.Id) === removedId);
              if (idx !== -1) page.splice(idx, 1);
            }
            return respond({ d: null });
          }
          // The flagged group's own membership, keyed off the BY-ID form of
          // the path so the 1.2 empty-group gate (which asks by name)
          // still reaches the adopted mock underneath and still sees empty.
          if (/sitegroups\(\d+\)\/users/.test(u)) {
            if (method === 'POST') {
              const added = JSON.parse(opts.body);
              if (!DROP_READBACK) {
                READER_MEMBER_PAGES[READER_MEMBER_PAGES.length - 1].push(
                  { Id: ENSURED.Id, Title: ENSURED.Title || '',
                    LoginName: added.LoginName });
              }
              // Somebody else's write landing in the same window.
              if (STRAY_ON_WRITE) {
                READER_MEMBER_PAGES[READER_MEMBER_PAGES.length - 1].push(
                  STRAY_ON_WRITE);
              }
              return respond({ d: { Id: ENSURED.Id, LoginName: added.LoginName } });
            }
            // Page 0 unless the caller followed a __next we handed out.
            // The follow-on URL keeps the sitegroups(N)/users shape so it
            // lands back here rather than falling through to the adopted
            // mock, which would answer an unrelated empty membership.
            const marked = /[?&]page=(\d+)/.exec(u);
            const page = marked ? Number(marked[1]) : 0;
            const payload = { d: { results: READER_MEMBER_PAGES[page] || [] } };
            if (page + 1 < READER_MEMBER_PAGES.length) {
              payload.d.__next =
                'https://example.sharepoint.com/_api/web/sitegroups(9)/users?page='
                + (page + 1);
            }
            return respond(payload);
          }
          return _beforeReader(url, opts);
        };
    """).replace(
        "__ENSURE_USER__", json.dumps(ensure_user),
    ).replace(
        "__MEMBER_PAGES__", json.dumps(pages),
    ).replace(
        "__DROP_READBACK__", "true" if drop_readback else "false",
    ).replace(
        "__STRAY_ON_WRITE__", json.dumps(stray_on_write),
    )


def _run_reader_deploy(
    ensure_user: dict[str, Any],
    *,
    members: list[dict[str, Any]] | None = None,
    member_pages: list[list[dict[str, Any]]] | None = None,
    drop_readback: bool = False,
    stray_on_write: dict[str, Any] | None = None,
) -> tuple[dict[str, Any], list[dict[str, Any]], str]:
    """Run the emitted deploy against the reader harness.

    Returns (summary, calls, output). The phase must actually have STARTED:
    a refusal test would otherwise pass against a run that aborted in 1.2
    and never reached the code under test at all.
    """
    script = _reader_harness(
        ensure_user, members=members, member_pages=member_pages,
        drop_readback=drop_readback, stray_on_write=stray_on_write,
    ) + "\n" + _reader_deploy_js().replace(
        "})();",
        "}))().then(r => { console.log('__RESULT__' + JSON.stringify(r));"
        " console.log('__CALLS__' + JSON.stringify(globalThis.__calls)); })",
    ).replace("(async () => {", "((async () => {", 1)
    output = _run(script)
    result_line = next(
        (ln for ln in output.splitlines() if ln.startswith("__RESULT__")), None,
    )
    assert result_line is not None, f"deploy.js did not return a summary:\n{output[-3000:]}"
    calls_line = next(
        (ln for ln in output.splitlines() if ln.startswith("__CALLS__")), None,
    )
    assert calls_line is not None, f"harness produced no call log:\n{output[-3000:]}"
    assert f"Starting Phase {pn('reader_enrolment')}" in output, (
        f"the reader-enrolment phase never ran:\n{output[-3000:]}"
    )
    return (
        json.loads(result_line.removeprefix("__RESULT__")),
        json.loads(calls_line.removeprefix("__CALLS__")),
        output,
    )


def _membership_writes(calls: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Every POST that adds somebody to a site group, by parsed body."""
    return [
        json.loads(c["body"]) for c in calls
        if c["method"] == "POST" and c.get("body") and _MEMBERSHIP_URL.search(c["url"])
    ]


def _reader_errors(summary: dict[str, Any]) -> list[Any]:
    return [
        err for err in (summary.get("errors") or [])
        if str(err.get("phase")) == pn("reader_enrolment")
    ]


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_security_group_is_refused_as_an_enterprise_reader() -> None:
    """Microsoft Learn: PrincipalType is None 0, User 1, DistributionList 2,
    SecurityGroup 4, SharePointGroup 8, and it carries [Flags], so the
    check is strict equality to 1, never a bitwise AND.

    `ensureuser` resolves a security group happily. Enrolling one would hand
    Read to everybody in it, and nothing downstream could tell: the deploy
    reads back byte-identical either way. Only the type differs from the
    success payload, so the type check is the only guard that can fire.
    """
    summary, calls, _ = _run_reader_deploy({
        **_RESOLVED_USER,
        "PrincipalType": 4,
        "LoginName": "c:0t.c|tenant|4d4a4d54-0b2e-4a1f-9b6c-2f0d7a0b1c3e",
        "Title": "Reporting Readers",
    })
    # The grant first: that is the damage, and the abort is only how the
    # script avoids it. Asserted in this order so removing the guard fails
    # on "a group was enrolled" rather than on a summary key.
    assert not _membership_writes(calls), (
        "a security group was enrolled: every member of it now holds Read"
    )
    assert summary.get("aborted") == "reader-enrolment-errors", summary
    assert _reader_errors(summary), summary


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_the_everyone_claim_is_refused_even_though_it_types_as_a_user() -> None:
    """`spo-grid-all-users` is the one mistake here with no cheap undo.

    On the one tenant this has been measured on (2026-08-12, group B of
    `test/manual/enterprise-reader-probe.js`) it came back typed 4, which
    the strict type check refuses by itself, so the needle is belt and
    braces behind that check, not the thing holding the door. This test
    hands it PrincipalType 1 anyway, because ONE TENANT IS ONE DATA POINT
    and the needle exists for the tenant that types it differently. That is
    also the only payload under which removing the needle can be watched
    failing.

    The payload keeps the matching Email deliberately, so neither the type
    check nor the identity check can be what refuses it. The claims check
    is on its own here, which is the only way removing it can be watched
    failing.
    """
    summary, calls, _ = _run_reader_deploy({
        **_RESOLVED_USER,
        "LoginName": "c:0-.f|rolemanager|spo-grid-all-users/contoso",
        "Title": "Everyone except external users",
    })
    assert not _membership_writes(calls), (
        "an everyone-claim was enrolled: every user in the tenant now holds Read"
    )
    assert summary.get("aborted") == "reader-enrolment-errors", summary


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_mismatched_identity_is_refused() -> None:
    """`ensureuser` resolving something other than what was asked for is the
    quiet failure: the deploy succeeds and the wrong account holds Read.

    A real user, correctly typed, with a real login. Only the address
    differs from the one the build asked for.
    """
    summary, calls, _ = _run_reader_deploy({
        **_RESOLVED_USER,
        "Id": 43,
        "LoginName": "i:0#.f|membership|someone-else@example.org",
        "Title": "Someone Else",
        "Email": "someone-else@example.org",
    })
    assert not _membership_writes(calls), (
        "an account other than the one asked for was enrolled"
    )
    assert summary.get("aborted") == "reader-enrolment-errors", summary


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_resolved_user_is_enrolled_and_the_membership_read_back() -> None:
    """The success path: the account IS added, by its resolved LoginName,
    and the run re-reads the membership afterwards rather than trusting the
    POST's own answer."""
    summary, calls, output = _run_reader_deploy(_RESOLVED_USER)
    assert not _reader_errors(summary), summary
    writes = _membership_writes(calls)
    assert [w["LoginName"] for w in writes] == [_RESOLVED_USER["LoginName"]], writes
    membership = [
        i for i, c in enumerate(calls) if _MEMBERSHIP_URL.search(c["url"])
    ]
    posted = next(
        i for i in membership
        if calls[i]["method"] == "POST" and calls[i].get("body")
    )
    assert any(
        i > posted and calls[i]["method"] == "GET" for i in membership
    ), f"the membership was never re-read after the write: {[calls[i] for i in membership]}"
    assert _RESOLVED_USER["Title"] in output


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_an_alias_mailbox_still_matches_the_requested_upn() -> None:
    """The identity check accepts the Email OR the LoginName's UPN part.

    An account whose mailbox address differs from its UPN is ordinary, and
    refusing it would send an operator looking for a fault that is not
    there. The account here is the right one (its claims login ends in the
    requested UPN), so it must be enrolled, not refused.
    """
    summary, calls, _ = _run_reader_deploy({
        **_RESOLVED_USER, "Email": "svc.reporting.alias@example.org",
    })
    assert not _reader_errors(summary), summary
    assert [w["LoginName"] for w in _membership_writes(calls)] == [
        _RESOLVED_USER["LoginName"],
    ]


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_membership_that_does_not_read_back_aborts() -> None:
    """The house rule: anything that writes reads back and verifies.

    SharePoint answering 200 is not evidence the membership exists. The
    harness accepts the POST and drops it, which is what a silently refused
    write looks like from the script's side. Reporting that as success is
    worse than failing, because the operator stops looking.
    """
    summary, calls, _ = _run_reader_deploy(_RESOLVED_USER, drop_readback=True)
    assert _membership_writes(calls), "the run never attempted the enrolment"
    assert summary.get("aborted") == "reader-enrolment-errors", summary


_OTHER_MEMBER: dict[str, Any] = {
    "Id": 7, "Title": "Data Team",
    "LoginName": "i:0#.f|membership|data-team@example.org",
}


def _removals(calls: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Every call that could take somebody OUT of a site group."""
    return [
        c for c in calls
        if "removebyid" in c["url"].lower() or "removebyloginname" in c["url"].lower()
        or c["method"] == "DELETE"
    ]


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_an_unexpected_member_aborts_the_run_and_is_never_removed() -> None:
    """A principal in the group that is not the named reader stops the run.

    This replaced an INFO line that let the run continue, and the case that
    forced the change is mundane: enrol a mistyped-but-valid address, notice,
    redeploy with the right one, and BOTH accounts hold Read on every list
    this bundle provisions, permanently, since nothing here removes anyone.
    The only trace was one INFO line in a run that reported success.

    Three things are asserted, and the ORDER matters. The damage is the
    grant, so "nothing was POSTed" comes first: deleting the gate must fail
    on a second reader having been enrolled, not on a summary key. Then the
    abort. Then, still, that nobody was removed. That half of the old
    behaviour is unchanged and this is the test pinning it. A gate that
    "fixed" the problem by evicting the stranger would pass the first two
    assertions and be a far worse tool.
    """
    summary, calls, output = _run_reader_deploy(
        _RESOLVED_USER, members=[_OTHER_MEMBER],
    )
    assert not _membership_writes(calls), (
        "a second account was enrolled into a group that already held "
        "somebody else; both now hold Read on every list in the bundle"
    )
    assert summary.get("aborted") == "reader-enrolment-errors", summary
    assert not _removals(calls), (
        f"the phase removed an existing member: {_removals(calls)}"
    )
    # Actionable: the operator has to be able to go and find the principal,
    # which needs the login name, not just a display title.
    errors = _reader_errors(summary)
    assert errors, summary
    message = str(errors[0]["error"])
    assert _OTHER_MEMBER["Title"] in message, message
    assert _OTHER_MEMBER["LoginName"] in message, message
    assert "Site permissions" in message, message
    assert "--enterprise-reader" in message, message
    assert _OTHER_MEMBER["Title"] in output


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_an_unexpected_member_on_a_later_page_still_aborts() -> None:
    """The gate reads every page of the membership, not the first one.

    SharePoint pages `sitegroups(N)/users` and hands back a `__next`. A gate
    that stopped at page one would be defeated by the group simply being
    big, and it would look like it worked, because the small groups every
    test uses fit in one page. Page one here is EMPTY and the stranger is
    alone on page two, so a first-page-only read sees a group with no
    members at all and enrols straight past them.
    """
    summary, calls, _ = _run_reader_deploy(_RESOLVED_USER, member_pages=[
        [],
        [_OTHER_MEMBER],
    ])
    assert not _membership_writes(calls), (
        "a member on page two was missed and the enrolment went ahead"
    )
    assert summary.get("aborted") == "reader-enrolment-errors", summary
    assert not _removals(calls)


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_the_named_reader_plus_a_stranger_still_aborts() -> None:
    """The named account already being a member does not excuse the other one.

    Ordering guard: the idempotence check ("already a member, skip") must
    not run before the gate, or the very redeploy that follows a mistyped
    address would sail through, which is the exact sequence this feature
    exists to catch.
    """
    summary, calls, _ = _run_reader_deploy(_RESOLVED_USER, members=[
        {"Id": _RESOLVED_USER["Id"], "Title": _RESOLVED_USER["Title"],
         "LoginName": _RESOLVED_USER["LoginName"]},
        _OTHER_MEMBER,
    ])
    assert not _membership_writes(calls)
    assert summary.get("aborted") == "reader-enrolment-errors", summary
    assert not _removals(calls)


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_principal_added_during_the_run_is_caught_by_the_read_back() -> None:
    """The gate above reads the state this run FOUND; this reads what it LEFT.

    The before-read gate cannot see a principal that arrives after it has
    already passed. Checking only that the reader is present made the
    read-back a presence check, so another administrator adding somebody
    between the two reads left the run reporting success on a group whose
    entire purpose is that it holds one account.

    A deploy is pasted into a site while people are working in it, so the
    window is not theoretical. `stray_on_write` models exactly that and
    nothing else: the membership is empty when the gate runs, and the
    foreign principal appears at the moment of the write.

    Nothing here removes the STRAY -- the same reason the before-read gives:
    membership is an operator-owned concern and this is a gate, not a
    reconciler for an account this run did not add. But (task 6,
    security-phase atomicity, #213 form 1) this run's OWN account no longer
    stays enrolled after this abort: deploy.js.j2's finally drains it,
    because the run never reached the end. This used to be the case that
    left BOTH accounts in two concurrent deploys permanently enrolled.
    """
    summary, calls, output = _run_reader_deploy(
        _RESOLVED_USER, members=[], stray_on_write=_OTHER_MEMBER,
    )
    # The enrolment really happened -- otherwise this would be re-testing
    # the before-read gate under a new name.
    assert _membership_writes(calls), output[-2000:]
    # The message assertion comes FIRST deliberately. Without the read-back
    # invariant the run carries on and aborts later for an unrelated reason,
    # so asserting the abort code first reports that later reason and buries
    # the defect. This one names it.
    errors = _reader_errors(summary)
    assert errors, summary
    assert "while this script was running" in str(errors), errors
    assert summary.get("aborted") == "reader-enrolment-errors", summary
    removals = _removals(calls)
    assert any(f"removebyid({_RESOLVED_USER['Id']})" in c["url"] for c in removals), (
        f"the reader this run just enrolled was left in place after the abort: {removals}"
    )
    assert not any(f"removebyid({_OTHER_MEMBER['Id']})" in c["url"] for c in removals), (
        f"the stray, which this run never added, was removed: {removals}"
    )


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_an_already_enrolled_reader_is_not_added_twice() -> None:
    """Idempotence: a redeploy must not POST a membership that is already
    there, and must not treat the existing one as a failure.

    The gate above counts principals that are not the named reader, so this
    is the case that says it counts them CORRECTLY: a re-run with the same
    flag has to stay green, or the feature is unusable after its first use.
    """
    summary, calls, _ = _run_reader_deploy(_RESOLVED_USER, members=[{
        "Id": _RESOLVED_USER["Id"], "Title": _RESOLVED_USER["Title"],
        "LoginName": _RESOLVED_USER["LoginName"],
    }])
    assert not _reader_errors(summary), summary
    # Not `aborted is None`: this harness's run stops later in Phase 1 for
    # reasons that have nothing to do with the reader. What must be true is
    # that the reader phase is not what stopped it.
    assert summary.get("aborted") != "reader-enrolment-errors", summary
    assert not _membership_writes(calls), "an existing membership was re-POSTed"


# Task 6 (security-phase atomicity, #213 form 1): the reader enrolment must
# clean up after a run that adds the account and then fails LATER, and must
# leave it alone on a run that reaches the end. Two concurrent deploys
# naming different reader addresses used to both add their account, both
# abort, and both leave their account enrolled forever.


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_later_phase_failure_drains_the_reader_this_run_added() -> None:
    """The general form of #213's fix: the reader phase itself succeeds
    cleanly, and a LATER phase aborts. The account this run just added must
    not survive that abort -- only a run that reaches the end may leave it.

    Plain `_run_reader_deploy` already aborts in Phase 2.1 on this fixture's
    Lookup column, for reasons that have nothing to do with the reader (the
    same fact `test_an_already_enrolled_reader_is_not_added_twice` notes
    above) -- exactly the shape this test needs: a clean 1.4 followed by a
    dirty later phase, with no bespoke harness required to get there.
    """
    summary, calls, output = _run_reader_deploy(_RESOLVED_USER)
    assert not _reader_errors(summary), (
        f"the reader phase itself failed, so this does not test a LATER "
        f"phase's abort: {summary}"
    )
    assert summary.get("aborted") not in (None, "reader-enrolment-errors"), (
        f"the run did not abort in a later phase: {summary}\n{output[-2000:]}"
    )
    assert _membership_writes(calls), "the reader was never enrolled in the first place"
    removals = _removals(calls)
    assert any(f"removebyid({_RESOLVED_USER['Id']})" in c["url"] for c in removals), (
        f"a later phase's abort must remove the reader this run just enrolled: {removals}"
    )


# `_declared_pack`'s minimal all-text schema is what `_ADOPTED_HARNESS` can
# drive all the way to a clean DATA-phase finish
# (`test_a_declared_run_completes_every_phase_cleanly`); the shipped
# `sharepoint-mapping(-with-reader).yaml` fixture's Lookup and formula
# columns are not modelled that completely (see the test above), so "a
# clean run leaves the reader enrolled" needs this schema, not that
# fixture, to reach the end at all.
_READER_DECLARED_SECTION = """
    groups:
      - name: "Enterprise Reader"
        description: "Read-only enrolment target for --enterprise-reader."
        owner_group: "Site Owners"
        allow_members_edit_membership: false
        allow_request_to_join_leave: false
        auto_accept_request_to_join_leave: false
        only_allow_members_view_membership: false
        enroll_enterprise_reader: true

    list_permissions:
      default:
        site_role: default
        break_inheritance: true
        reconcile: exact
        assignments:
          - principal: { kind: group, name: "Enterprise Reader" }
            level: "Read"
"""


def _declared_reader_deploy_js(tmp_path: Path) -> str:
    """`_declared_deploy_js`, plus an enterprise-reader group with a Read
    ACL assignment, built with `--enterprise-reader`."""
    from dbml_sharepoint.generators.jsgen import generate_deploy_js
    from dbml_sharepoint.model.release import load_release

    schema, bundle = _declared_pack(tmp_path, _READER_DECLARED_SECTION)
    return generate_deploy_js(
        schema=schema,
        bundle=bundle,
        release=load_release(FIXTURES / "release.yaml"),
        site_url="https://example.sharepoint.com/sites/test",
        site_role="default",
        source_dbml="s.dbml",
        source_mtime="2026-05-04T00:00:00Z",
        generated_at="2026-05-04T00:00:00Z",
        enterprise_reader=_READER_ADDRESS,
    )


# The ACL phase resolves a declared assignment's level by name through
# `web/roledefinitions/getbyname`, the same surface `ROLE_DEF_STATE` already
# models for this tool's own custom levels. No existing test names a
# BUILT-IN level ('Read') in an ACL assignment, so `_ADOPTED_HARNESS` never
# needed to seed one; this is the first run to carry an enterprise-reader
# group's Read grant all the way through Phase 4.2.
_READER_ACL_HARNESS = _ADOPTED_HARNESS.replace(
    "'Schema Manager': {",
    "'Read': { Id: 100, Description: '', BasePermissions: { High: '0', Low: '138612833' } },\n"
    "      'Schema Manager': {",
)


def _reader_harness_for_declared_run(ensure_user: dict[str, Any]) -> str:
    """`_reader_harness`, rebuilt on `_READER_ACL_HARNESS` instead of the
    plain `_ADOPTED_HARNESS` it always starts from."""
    overlay = _reader_harness(ensure_user)[len(_ADOPTED_HARNESS):]
    return _READER_ACL_HARNESS + overlay


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_clean_end_to_end_run_leaves_the_reader_enrolled(tmp_path: Path) -> None:
    """The other half of #213's fix: a run that reaches the end must NOT
    remove the reader it just added. That membership is meant to outlive
    the run, and only a run that never gets here should undo it.
    """
    js = _declared_reader_deploy_js(tmp_path)
    script = _reader_harness_for_declared_run(_RESOLVED_USER) + "\n" + js.replace(
        "})();",
        "}))().then(r => { console.log('__RESULT__' + JSON.stringify(r));"
        " console.log('__CALLS__' + JSON.stringify(globalThis.__calls)); })",
    ).replace("(async () => {", "((async () => {", 1)
    output = _run(script)
    result_line = next(
        (ln for ln in output.splitlines() if ln.startswith("__RESULT__")), None,
    )
    assert result_line is not None, f"deploy.js did not return a summary:\n{output[-3000:]}"
    summary = json.loads(result_line.removeprefix("__RESULT__"))
    calls_line = next(
        (ln for ln in output.splitlines() if ln.startswith("__CALLS__")), None,
    )
    assert calls_line is not None, f"harness produced no call log:\n{output[-3000:]}"
    calls = json.loads(calls_line.removeprefix("__CALLS__"))

    assert summary.get("aborted") is None, summary
    assert summary.get("errors") == [], summary["errors"]
    assert _membership_writes(calls), "the reader was never enrolled"
    assert not _removals(calls), (
        f"a successful run removed the reader it just enrolled: {_removals(calls)}"
    )


# #209: a hand-made site group whose name happens to match a declared one
# used to be adopted by name alone, and the ACL phase then granted it
# whatever the mapping declares for that group. `_group_gate_deploy` lets
# these tests control ONE declared group's Description and paginated
# membership against `_ADOPTED_HARNESS`, leaving every other group at the
# shared defaults (unmarked, empty) so it is neither created nor refused.


def _run_group_verify_deploy(
    deploy_js: str, harness: str,
) -> tuple[dict[str, Any], list[dict[str, Any]], str]:
    """Run `deploy_js` against `harness` and return (summary, calls, output).

    Phase 'security' must actually have started, or a refusal assertion
    would pass against a run that never reached the group loop at all.
    """
    script = harness + "\n" + deploy_js.replace(
        "})();",
        "}))().then(r => { console.log('__RESULT__' + JSON.stringify(r));"
        " console.log('__CALLS__' + JSON.stringify(globalThis.__calls)); })",
    ).replace("(async () => {", "((async () => {", 1)
    output = _run(script)
    result_line = next(
        (ln for ln in output.splitlines() if ln.startswith("__RESULT__")), None,
    )
    assert result_line is not None, f"deploy.js did not return a summary:\n{output[-3000:]}"
    calls_line = next(
        (ln for ln in output.splitlines() if ln.startswith("__CALLS__")), None,
    )
    assert calls_line is not None, f"harness produced no call log:\n{output[-3000:]}"
    assert f"Starting Phase {pn('security')}" in output, (
        f"the security phase never ran:\n{output[-3000:]}"
    )
    return (
        json.loads(result_line.removeprefix("__RESULT__")),
        json.loads(calls_line.removeprefix("__CALLS__")),
        output,
    )


def _group_gate_deploy(
    deploy_js: str,
    group_name: str,
    *,
    description: str,
    member_pages: list[list[dict[str, Any]]],
) -> tuple[dict[str, Any], list[dict[str, Any]], str]:
    """Run `deploy_js` against `_ADOPTED_HARNESS` with one group's Description
    and membership overridden.
    """
    harness = _ADOPTED_HARNESS.replace(
        "const GROUP_DESCRIPTIONS = {};",
        f"const GROUP_DESCRIPTIONS = {json.dumps({group_name: description})};",
    ).replace(
        "const GROUP_MEMBER_PAGES = {};",
        f"const GROUP_MEMBER_PAGES = {json.dumps({group_name: member_pages})};",
    ).replace(
        "const KNOWN_GROUP_NAMES = [];",
        f"const KNOWN_GROUP_NAMES = {json.dumps([group_name])};",
    )
    return _run_group_verify_deploy(deploy_js, harness)


def _group_settings_writes(
    calls: list[dict[str, Any]], group_name: str,
) -> list[dict[str, Any]]:
    """Every POST that MERGEs settings, description included, onto the named
    group object itself, not its membership."""
    encoded = quote(group_name, safe="")
    return [
        c for c in calls
        if c["method"] == "POST" and c["body"]
        and f"sitegroups/getbyname('{encoded}')" in c["url"]
        and "/users" not in c["url"]
    ]


def _security_errors(summary: dict[str, Any]) -> list[Any]:
    return [
        err for err in (summary.get("errors") or [])
        if str(err.get("phase")) == pn("security")
    ]


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_an_unmarked_group_with_members_is_refused() -> None:
    """#209: adopting it would grant those members whatever the family
    declares. 'List Maintainer' is granted 'Schema Manager' in the plain
    fixture mapping, which can create and manage every list in the family."""
    summary, calls, _ = _group_gate_deploy(
        _deploy_js(), "List Maintainer",
        description="Our own group", member_pages=[[{"Id": 501}]],
    )
    # The grant is the damage; the refusal is only how the script avoids it.
    # Asserted first so removing the gate fails on "the group was
    # reconciled" rather than on a summary key.
    assert not _group_settings_writes(calls, "List Maintainer"), (
        "an unmarked group with members was reconciled before the refusal: "
        "its description and membership controls were rewritten"
    )
    errors = _security_errors(summary)
    assert errors, summary
    message = str(errors[0]["error"])
    assert "List Maintainer" in message, message
    assert "carries no" in message, message
    assert summary.get("aborted") == "phase-0-security-errors", summary


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_an_unmarked_but_empty_group_is_adopted_and_stamped() -> None:
    """A group nobody has joined yet carries no access to hand out, so it is
    adopted like any other pre-existing group and stamped with the marker
    that lets a later redeploy recognise it as this tool's own."""
    summary, calls, _ = _group_gate_deploy(
        _deploy_js(), "List Maintainer",
        description="Our own group", member_pages=[[]],
    )
    assert not _security_errors(summary), summary
    writes = _group_settings_writes(calls, "List Maintainer")
    assert writes, "an unmarked, empty group was never adopted"
    assert "Provisioned by dbml-sharepoint" in writes[0]["body"]


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_an_unmarked_group_with_a_member_on_a_later_page_is_refused() -> None:
    """`countGroupMembers` must follow every page, not just the first.

    The first page is empty and the only member sits on page two. A count
    that stopped after page one would read zero, adopt the group, and MERGE
    the family's grant onto it. This is the only guard surface on this
    branch with no coverage of its own `__next` pagination, so it gets a
    test that a broken loop cannot pass by accident:
    `test_an_unmarked_group_with_members_is_refused` puts every member on
    page one and would stay green even if pagination were deleted entirely.
    """
    summary, calls, _ = _group_gate_deploy(
        _deploy_js(), "List Maintainer",
        description="Our own group", member_pages=[[], [{"Id": 501}]],
    )
    assert not _group_settings_writes(calls, "List Maintainer"), (
        "an unmarked group with a member on a later page was reconciled "
        "before the refusal"
    )
    assert _security_errors(summary), summary


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_marked_group_with_members_is_adopted_silently() -> None:
    """A redeploy must not trip over the enterprise reader a prior run
    already enrolled into this same group. `enterprise_reader=None` keeps
    the reader-enrolment phase itself out of the emitted script, so this
    exercises only the adoption gate in the security phase that reconciles
    'Enterprise Reader' regardless of --enterprise-reader."""
    summary, calls, _ = _group_gate_deploy(
        _reader_deploy_js(enterprise_reader=None), "Enterprise Reader",
        description="Read-only accounts. "
            "Provisioned by dbml-sharepoint from simple-test for group Enterprise Reader.",
        member_pages=[[{"Id": 501}]],
    )
    assert not _security_errors(summary), summary
    assert _group_settings_writes(calls, "Enterprise Reader"), (
        "a marked group's settings were never reconciled"
    )


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_group_marked_by_another_family_with_members_is_refused() -> None:
    """The gate must compare the exact marker this declaration expects, not
    the shared prefix every family's marker starts with. A group another
    family stamped and populated satisfied the old prefix-only test, and the
    ACL phase then granted those members whatever THIS family declares."""
    summary, calls, _ = _group_gate_deploy(
        _reader_deploy_js(enterprise_reader=None), "Enterprise Reader",
        description="Read-only accounts. "
            "Provisioned by dbml-sharepoint from other-family for group Enterprise Reader.",
        member_pages=[[{"Id": 501}]],
    )
    assert not _group_settings_writes(calls, "Enterprise Reader"), (
        "a group marked by another family was reconciled before the refusal"
    )
    errors = _security_errors(summary)
    assert errors, summary
    message = str(errors[0]["error"])
    assert "Enterprise Reader" in message, message
    assert "carries no" in message, message
    assert summary.get("aborted") == "phase-0-security-errors", summary


# Task 7: `mergeResp.ok` and the create POST's `ok` only say the tenant
# accepted the request, not that it stored what was sent. `verifyGroupSettings`
# reads the group back after both the create and the reconcile write and
# compares every field it wrote. `GROUP_DROP_FIELD_ON_WRITE` and
# `GROUP_COERCE_AUTO_ACCEPT` model the two ways the mock, like the tenant, can
# answer 200 while storing something other than what was sent.


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_group_description_the_tenant_did_not_store_fails_closed() -> None:
    """AGENTS.md: anything that writes must read back and verify."""
    harness = _ADOPTED_HARNESS.replace(
        "const GROUP_DROP_FIELD_ON_WRITE = null;",
        "const GROUP_DROP_FIELD_ON_WRITE = 'Description';",
    )
    summary, _, _ = _run_group_verify_deploy(_deploy_js(), harness)
    errors = _security_errors(summary)
    assert errors, summary
    message = str(errors[0]["error"])
    assert "List Maintainer" in message, message
    assert "Description" in message, message
    assert summary.get("aborted") == "phase-0-security-errors", summary


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_flag_the_tenant_ignored_fails_closed() -> None:
    """OnlyAllowMembersViewMembership is part of the security boundary.

    The fixture declares it false, which is also the mock's untouched
    default, so dropping the write would leave the state unchanged and
    prove nothing. The schema text is overridden to true so the drop is
    observable: the state stays at the untouched default instead of
    picking up what was sent.
    """
    js = _deploy_js().replace(
        '"only_allow_members_view_membership": false',
        '"only_allow_members_view_membership": true', 1,
    )
    assert '"only_allow_members_view_membership": true' in js
    harness = _ADOPTED_HARNESS.replace(
        "const GROUP_DROP_FIELD_ON_WRITE = null;",
        "const GROUP_DROP_FIELD_ON_WRITE = 'OnlyAllowMembersViewMembership';",
    )
    summary, _, _ = _run_group_verify_deploy(js, harness)
    errors = _security_errors(summary)
    assert errors, summary
    message = str(errors[0]["error"])
    assert "List Maintainer" in message, message
    assert "OnlyAllowMembersViewMembership" in message, message
    assert summary.get("aborted") == "phase-0-security-errors", summary


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_auto_accept_is_compared_against_the_coerced_value() -> None:
    """group-description-probe.js, G9/G10 (2026-08-13 and 2026-08-14): the
    tenant stores AutoAcceptRequestToJoinLeave as false whenever the written
    AllowRequestToJoinLeave is false, no matter what was sent for AutoAccept.
    `group_auto_accept_without_requests` refuses a mapping that declares
    that pair, so no shipped mapping can reach this branch through the CLI.
    The schema text is overridden directly to reach it anyway, the same way
    every other test in this file calls `generate_deploy_js` without going
    through the build-time checks.

    Comparing the read-back against the value SENT, rather than the coerced
    one, would fail here: SENT is true, the tenant stores false, and the
    deploy must accept that as correct rather than abort. Getting this wrong
    aborts a redeploy for every shipped family, since every one of them
    declares AllowRequestToJoinLeave false.
    """
    js = _deploy_js().replace(
        '"auto_accept_request_to_join_leave": false',
        '"auto_accept_request_to_join_leave": true', 1,
    )
    assert '"auto_accept_request_to_join_leave": true' in js
    harness = _ADOPTED_HARNESS.replace(
        "const GROUP_COERCE_AUTO_ACCEPT = false;",
        "const GROUP_COERCE_AUTO_ACCEPT = true;",
    )
    summary, _, _ = _run_group_verify_deploy(js, harness)
    assert not _security_errors(summary), summary


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_reconciled_group_setting_the_tenant_did_not_store_fails_closed() -> None:
    """The three tests above all exercise the create path (`KNOWN_GROUP_NAMES`
    empty). `verifyGroupSettings` is called on the reconcile path too, and
    that call has no coverage of its own without this: `KNOWN_GROUP_NAMES`
    names the group, and its pre-existing Description already carries the
    marker (`Stale note...`), so this takes the adopt-and-reconcile branch,
    never the create branch, and the drop can only be caught by the
    reconcile read-back.
    """
    group_name = "List Maintainer"
    stale_description = (
        "Stale note. "
        "Provisioned by dbml-sharepoint from simple-test for group List Maintainer."
    )
    harness = _ADOPTED_HARNESS.replace(
        "const GROUP_DESCRIPTIONS = {};",
        f"const GROUP_DESCRIPTIONS = {json.dumps({group_name: stale_description})};",
    ).replace(
        "const GROUP_MEMBER_PAGES = {};",
        f"const GROUP_MEMBER_PAGES = {json.dumps({group_name: [[]]})};",
    ).replace(
        "const KNOWN_GROUP_NAMES = [];",
        f"const KNOWN_GROUP_NAMES = {json.dumps([group_name])};",
    ).replace(
        "const GROUP_DROP_FIELD_ON_WRITE = null;",
        "const GROUP_DROP_FIELD_ON_WRITE = 'Description';",
    )
    summary, calls, _ = _run_group_verify_deploy(_deploy_js(), harness)
    assert _group_settings_writes(calls, group_name), (
        "the reconcile MERGE never happened"
    )
    errors = _security_errors(summary)
    assert errors, summary
    assert "group" in errors[0], errors[0]
    message = str(errors[0]["error"])
    assert group_name in message, message
    assert "Description" in message, message
    assert summary.get("aborted") == "phase-0-security-errors", summary


# Every survey in the security phase now runs before every create, so an
# adopt decision's owner resolve, which reads a SECOND group (the declared
# owner_group), can hit a custom group this SAME pass has not created yet.
# `_owner_pending_groups_deploy_js` splits the fixture's one declared group
# into two so the adopted one names the about-to-be-created one as its owner.


def _owner_pending_groups_deploy_js() -> str:
    """`_deploy_js()` with the fixture's one declared group ('List
    Maintainer') split into two: it now declares owner_group 'Group B', a
    second custom group this same declaration also creates. Mutates the
    generated JSON directly, the same way `test_auto_accept_is_compared_...`
    does, rather than adding a second group to the shared mapping fixture.
    """
    js = _deploy_js()
    match = re.search(r'"groups": (\[.*?\n  \])', js, re.DOTALL)
    assert match, "groups array not found in generated deploy.js"
    groups = json.loads(match.group(1))
    assert len(groups) == 1, groups
    list_maintainer = dict(groups[0])
    assert list_maintainer["owner_group"] == "Site Owners", list_maintainer
    list_maintainer["owner_group"] = "Group B"
    group_b = dict(list_maintainer)
    group_b["name"] = "Group B"
    group_b["description"] = "Group B."
    group_b["owner_group"] = "Site Owners"
    group_b["require_empty_at_deploy"] = False
    new_groups = json.dumps([group_b, list_maintainer], indent=2).replace("\n", "\n  ")
    return js[: match.start(1)] + new_groups + js[match.end(1):]


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_an_adopted_group_owned_by_a_group_pending_creation_still_deploys() -> None:
    """'List Maintainer' already exists and declares owner_group 'Group B',
    which is declared but absent, so this same pass decides to create it.
    Resolving 'List Maintainer's owner during the survey, before Group B
    exists, would 404 and abort the whole phase; the fix defers that resolve
    to applyGroupDecision, which runs after Group B's own create has
    applied. Verified by mutation: forcing the resolve back into the survey
    unconditionally reproduces the abort this test would otherwise miss.
    """
    js = _owner_pending_groups_deploy_js()
    harness = _ADOPTED_HARNESS.replace(
        "const GROUP_DESCRIPTIONS = {};",
        f"const GROUP_DESCRIPTIONS = {json.dumps({
            'List Maintainer': 'Test group. Provisioned by dbml-sharepoint from simple-test.',
        })};",
    ).replace(
        "const GROUP_MEMBER_PAGES = {};",
        f"const GROUP_MEMBER_PAGES = {json.dumps({'List Maintainer': [[]]})};",
    ).replace(
        "const KNOWN_GROUP_NAMES = [];",
        f"const KNOWN_GROUP_NAMES = {json.dumps(['List Maintainer'])};",
    ).replace(
        "const GROUP_IDS = {};",
        f"const GROUP_IDS = {json.dumps({'List Maintainer': 101, 'Group B': 102})};",
    ).replace(
        "const GROUP_CURRENT_OWNER = { 9: { Id: 3, Title: 'Site Owners', PrincipalType: 8 } };",
        "const GROUP_CURRENT_OWNER = { 9: { Id: 3, Title: 'Site Owners', PrincipalType: 8 }, "
        "101: { Id: 102, Title: 'Group B', PrincipalType: 8 } };",
    )
    summary, calls, output = _run_group_verify_deploy(js, harness)
    assert not _security_errors(summary), summary
    assert summary.get("aborted") != "phase-0-security-errors", summary
    create_indices = [
        i for i, c in enumerate(calls)
        if c["method"] == "POST" and c["url"].endswith("/sitegroups") and c["body"]
        and json.loads(c["body"]).get("Title") == "Group B"
    ]
    assert create_indices, f"Group B was never created:\n{output[-3000:]}"
    owner_resolve_indices = [
        i for i, c in enumerate(calls)
        if c["method"] == "GET"
        and "sitegroups/getbyname('Group%20B')" in c["url"]
    ]
    assert owner_resolve_indices, f"'List Maintainer's owner was never resolved:\n{output[-3000:]}"
    assert min(owner_resolve_indices) > create_indices[0], (
        "the owner resolve for 'List Maintainer' ran before Group B was created: "
        f"resolve at {owner_resolve_indices}, create at {create_indices[0]}"
    )


# #224: `_security_principals.js.j2` adopted any role definition whose name
# matched a declared one and MERGEd the declared bitmap onto it. A role
# definition is SITE-SCOPED, so a hand-made level sharing a declared name and
# assigned on lists this tool never reads had its bitmap silently
# overwritten. `_role_def_gate_deploy` lets these tests control the
# fixture's one declared level, 'Schema Manager', against `_ADOPTED_HARNESS`.


def _role_def_gate_deploy(
    deploy_js: str,
    *,
    absent: bool = False,
    description_override: str | None = None,
) -> tuple[dict[str, Any], list[dict[str, Any]], str]:
    """Run `deploy_js` against `_ADOPTED_HARNESS` with the fixture's one
    declared permission level, 'Schema Manager', either absent (so the
    CREATE path runs) or present with its Description overridden (so the
    #224 adoption gate can be exercised against an unmarked level, or one
    another family stamped).
    """
    harness = _ADOPTED_HARNESS.replace(
        "const ROLE_DEF_ABSENT = false;",
        f"const ROLE_DEF_ABSENT = {json.dumps(absent)};",
    ).replace(
        "const ROLE_DEF_DESCRIPTION_OVERRIDE = null;",
        f"const ROLE_DEF_DESCRIPTION_OVERRIDE = {json.dumps(description_override)};",
    )
    return _run_group_verify_deploy(deploy_js, harness)


def _role_def_merge_writes(
    calls: list[dict[str, Any]], level_name: str,
) -> list[dict[str, Any]]:
    """Every POST that MERGEs settings onto the named role definition itself."""
    encoded = quote(level_name, safe="")
    return [
        c for c in calls
        if c["method"] == "POST" and c["body"]
        and f"roledefinitions/getbyname('{encoded}')" in c["url"]
    ]


def _role_def_create_writes(calls: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Every POST that creates a new role definition."""
    return [
        c for c in calls
        if c["method"] == "POST" and c["body"]
        and c["url"].endswith("/web/roledefinitions")
    ]


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_an_unmarked_permission_level_is_refused() -> None:
    """MARKER ONLY, unlike the group gate: no usage count can clear this
    refusal. Default `_ADOPTED_HARNESS` state reports zero web-scope role
    assignments for it, and it is still refused. `_acls.js.j2` assigns a
    permission level at LIST scope, which a web-scope count cannot see, so
    treating an unmeasured surface as empty would adopt exactly the level
    #224 exists to stop adopting."""
    summary, calls, _ = _role_def_gate_deploy(
        _deploy_js(), description_override="Our own level.",
    )
    # The overwritten bitmap is the damage; the refusal is only how the
    # script avoids it. Asserted first so removing the gate fails on "the
    # level was reconciled" rather than on a summary key.
    assert not _role_def_merge_writes(calls, "Schema Manager"), (
        "an unmarked permission level was reconciled before the refusal: "
        "its Description and BasePermissions were rewritten"
    )
    errors = _security_errors(summary)
    assert errors, summary
    message = str(errors[0]["error"])
    assert "Schema Manager" in message, message
    assert "carries no" in message, message
    assert summary.get("aborted") == "phase-0-security-errors", summary


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_permission_level_marked_by_another_family_is_refused() -> None:
    """The gate compares the exact marker THIS declaration expects, not the
    shared 'Provisioned by dbml-sharepoint' prefix every family's marker
    starts with, so a level another family stamped cannot satisfy it."""
    summary, calls, _ = _role_def_gate_deploy(
        _deploy_js(),
        description_override="Provisioned by dbml-sharepoint from other-family.",
    )
    assert not _role_def_merge_writes(calls, "Schema Manager"), (
        "a permission level marked by another family was reconciled before the refusal"
    )
    errors = _security_errors(summary)
    assert errors, summary
    message = str(errors[0]["error"])
    assert "Schema Manager" in message, message
    assert "carries no" in message, message
    assert summary.get("aborted") == "phase-0-security-errors", summary


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_marked_permission_level_is_adopted_silently() -> None:
    """The default `_ADOPTED_HARNESS` state already carries this
    declaration's marker, matching a prior run of the same family: the
    level is reconciled without a security error."""
    summary, calls, _ = _role_def_gate_deploy(_deploy_js())
    assert not _security_errors(summary), summary
    assert _role_def_merge_writes(calls, "Schema Manager"), (
        "a marked permission level's settings were never reconciled"
    )


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_permission_level_created_fresh_is_stamped() -> None:
    """A level this deploy creates must carry the marker that lets a later
    redeploy recognise it as this tool's own."""
    summary, calls, _ = _role_def_gate_deploy(_deploy_js(), absent=True)
    assert not _security_errors(summary), summary
    writes = _role_def_create_writes(calls)
    assert writes, "a fresh permission level was never created"
    marker = "Provisioned by dbml-sharepoint from simple-test for level Schema Manager."
    assert marker in writes[0]["body"]


# Task 5 (#224): `mergeResp.ok` and the create POST's `ok` only say the
# tenant accepted the request, not that it stored what was sent.
# `verifyLevelSettings` reads the level back after both the create and the
# MERGE and compares Description and both bitmap halves.
# `ROLE_DEF_DROP_FIELD_ON_WRITE` models a write the tenant 200s and
# discards, the same way `GROUP_DROP_FIELD_ON_WRITE` does for a site group.


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_permission_level_description_the_tenant_did_not_store_fails_closed() -> None:
    """AGENTS.md: anything that writes must read back and verify.

    The stale description still carries this family's marker, so the run
    takes the adopt-and-reconcile branch rather than the refusal gate; only
    the read-back after the MERGE can catch the drop.
    """
    stale_description = (
        "Stale note. "
        "Provisioned by dbml-sharepoint from simple-test for level Schema Manager."
    )
    harness = _ADOPTED_HARNESS.replace(
        "const ROLE_DEF_DESCRIPTION_OVERRIDE = null;",
        f"const ROLE_DEF_DESCRIPTION_OVERRIDE = {json.dumps(stale_description)};",
    ).replace(
        "const ROLE_DEF_DROP_FIELD_ON_WRITE = null;",
        "const ROLE_DEF_DROP_FIELD_ON_WRITE = 'Description';",
    )
    summary, calls, _ = _run_group_verify_deploy(_deploy_js(), harness)
    assert _role_def_merge_writes(calls, "Schema Manager"), (
        "the reconcile MERGE never happened"
    )
    errors = _security_errors(summary)
    assert errors, summary
    message = str(errors[0]["error"])
    assert "Schema Manager" in message, message
    assert "Description" in message, message
    assert summary.get("aborted") == "phase-0-security-errors", summary


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_permission_level_base_permissions_the_tenant_did_not_store_fails_closed() -> None:
    """A dropped bitmap half is the exact failure #224 exists to catch: the
    MERGE reports success while the level keeps its old permissions. The
    declared Low is overridden so the drop is observable, since the mock's
    stored default otherwise already equals the undisturbed declared value.
    """
    js = _deploy_js().replace('"low": "2049"', '"low": "4098"', 1)
    assert '"low": "4098"' in js
    harness = _ADOPTED_HARNESS.replace(
        "const ROLE_DEF_DROP_FIELD_ON_WRITE = null;",
        "const ROLE_DEF_DROP_FIELD_ON_WRITE = 'Low';",
    )
    summary, calls, _ = _run_group_verify_deploy(js, harness)
    assert _role_def_merge_writes(calls, "Schema Manager"), (
        "the reconcile MERGE never happened"
    )
    errors = _security_errors(summary)
    assert errors, summary
    assert "permissionLevel" in errors[0], errors[0]
    message = str(errors[0]["error"])
    assert "Schema Manager" in message, message
    assert "BasePermissions" in message, message
    assert summary.get("aborted") == "phase-0-security-errors", summary


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_freshly_created_level_base_permissions_the_tenant_did_not_store_fails_closed() -> None:
    """The two drop-field tests above only drive the MERGE (adopt) branch.
    `verifyLevelSettings` is called separately after CREATE, and deleting
    that call left every other test in this file green: the create-body
    assertion in `test_a_permission_level_created_fresh_is_stamped` does not
    move when the post-create read-back is skipped.

    `roleDefState`'s untouched default for a never-seen name already has
    BasePermissions.Low '0', which differs from the fixture's declared
    '2049' on its own, so no override of the declared value is needed here
    to make the drop observable, unlike the MERGE-path test above it.
    """
    harness = _ADOPTED_HARNESS.replace(
        "const ROLE_DEF_ABSENT = false;",
        "const ROLE_DEF_ABSENT = true;",
    ).replace(
        "const ROLE_DEF_DROP_FIELD_ON_WRITE = null;",
        "const ROLE_DEF_DROP_FIELD_ON_WRITE = 'Low';",
    )
    summary, calls, _ = _run_group_verify_deploy(_deploy_js(), harness)
    assert _role_def_create_writes(calls), "the create POST never happened"
    errors = _security_errors(summary)
    assert errors, summary
    message = str(errors[0]["error"])
    assert "Schema Manager" in message, message
    assert "BasePermissions" in message, message
    assert summary.get("aborted") == "phase-0-security-errors", summary


# Task 4 (security-phase-atomicity): the abort check used to run only after
# BOTH loops (levels, then groups) had finished, so a refusal on the first
# object did not stop a write on a LATER one. `_security_writes` names every
# POST the phase can issue, regardless of which object it belongs to, so the
# tests below anchor on the absence of writes in the call log rather than on
# a summary key, matching AGENTS.md's evidence rule: the gate must never be
# softenable without a test failing on the write itself.


def _security_writes(calls: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Every POST `_security_principals.js.j2` can issue: a permission-level
    create or MERGE, a site-group create or MERGE, or the CSOM ProcessQuery
    owner correction. All three only ever fire from the apply loop."""
    return [
        c for c in calls
        if c["method"] == "POST"
        and (
            "sitegroups" in c["url"]
            or "roledefinitions" in c["url"]
            or "ProcessQuery" in c["url"]
        )
    ]


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_the_apply_pass_takes_its_own_fresh_digest() -> None:
    """digest0 is captured near the top of phase 1.2, before the whole
    survey (every level probe, the group enumeration, every adopt-path
    owner read and membership count) now runs ahead of it, where before this
    task the first write followed the fetch almost immediately. A create
    write that reused digest0 directly, without asking getDigest() again,
    would carry whatever was fetched before the survey started rather than
    a digest taken right before the apply pass's first write.

    getDigest() caches for at least 60s (`_digest_cached.js.j2`), which a
    synchronous test cannot outlast, so the cache guard is disabled here to
    make every call a real fetch: the count of `contextinfo` POSTs before
    the first write is then a reliable proxy for whether the apply pass took
    its own fresh digest, rather than reusing the one taken before survey.
    """
    js = _deploy_js().replace(
        "if (cachedDigest && Date.now() < digestExpiresAt) return cachedDigest;",
        "if (false) return cachedDigest; // test: force every call to re-fetch",
    )
    assert "if (false) return cachedDigest" in js, "getDigest cache guard not found"
    summary, calls, output = _role_def_gate_deploy(js, absent=True)
    assert not _security_errors(summary), summary
    creates = _role_def_create_writes(calls)
    assert creates, f"a fresh permission level was never created:\n{output[-3000:]}"
    first_write_index = calls.index(creates[0])
    digests_before_first_write = [
        c for c in calls[:first_write_index] if "contextinfo" in c["url"]
    ]
    assert len(digests_before_first_write) >= 2, (
        "the apply pass reused the digest fetched before the survey instead of "
        f"taking its own fresh one: {digests_before_first_write}"
    )


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_refused_level_blocks_the_group_create_that_used_to_follow_it() -> None:
    """Before this task, the level loop ran survey-then-apply per object and
    only checked `summary.errors` after BOTH loops. A refused level did not
    stop the group loop a few lines later from creating 'List Maintainer'.
    """
    summary, calls, output = _role_def_gate_deploy(
        _deploy_js(), description_override="Our own level.",
    )
    assert not _security_writes(calls), (
        f"a refused permission level did not stop a write on another "
        f"object\n{output[-2000:]}"
    )
    errors = _security_errors(summary)
    assert errors, summary
    assert summary.get("aborted") == "phase-0-security-errors", summary


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_refused_group_blocks_the_permission_level_reconcile_too() -> None:
    """Symmetric case: with no override the fixture's one declared level
    ('Schema Manager') is adopted and reconciled cleanly on its own
    (`test_a_marked_permission_level_is_adopted_silently`). Refusing the
    group here must stop that reconcile from happening, whichever loop ran
    first.
    """
    summary, calls, output = _group_gate_deploy(
        _deploy_js(), "List Maintainer",
        description="Our own group", member_pages=[[{"Id": 501}]],
    )
    assert not _security_writes(calls), (
        f"a refused site group did not stop the permission level reconcile "
        f"that would otherwise have run\n{output[-2000:]}"
    )
    errors = _security_errors(summary)
    assert errors, summary
    assert summary.get("aborted") == "phase-0-security-errors", summary


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_two_refusals_both_appear_in_the_transcript() -> None:
    """Surveying must not short-circuit on the first refusal: an operator
    who fixes one blocker and redeploys must not just meet the next one
    they were never told about."""
    harness = _ADOPTED_HARNESS.replace(
        "const ROLE_DEF_DESCRIPTION_OVERRIDE = null;",
        f"const ROLE_DEF_DESCRIPTION_OVERRIDE = {json.dumps('Our own level.')};",
    ).replace(
        "const GROUP_DESCRIPTIONS = {};",
        f"const GROUP_DESCRIPTIONS = {json.dumps({'List Maintainer': 'Our own group'})};",
    ).replace(
        "const GROUP_MEMBER_PAGES = {};",
        f"const GROUP_MEMBER_PAGES = {json.dumps({'List Maintainer': [[{'Id': 501}]]})};",
    ).replace(
        "const KNOWN_GROUP_NAMES = [];",
        f"const KNOWN_GROUP_NAMES = {json.dumps(['List Maintainer'])};",
    )
    summary, calls, output = _run_group_verify_deploy(_deploy_js(), harness)
    assert not _security_writes(calls), output[-2000:]
    errors = _security_errors(summary)
    messages = [str(e["error"]) for e in errors]
    assert any("Schema Manager" in m for m in messages), errors
    assert any("List Maintainer" in m for m in messages), errors
    assert len(errors) == 2, (
        f"only one of two refusals reached the transcript: {errors}"
    )
    assert summary.get("aborted") == "phase-0-security-errors", summary


# A genuine survey FAILURE (not a refusal): the permission-level existence
# probe answers a real HTTP error rather than a filtered result set.
# `surveyLevel` throws in that case, and the per-object catch around it must
# still turn that into the same structured summary a refusal produces,
# rather than letting it escape the phase, the `try` in deploy.js.j2, and
# the async IIFE as an unhandled rejection -- which the harness would
# surface as a missing `__RESULT__` line, since nothing would ever call the
# `.then()` that prints it.
_SURVEY_FAILURE_HARNESS = _ADOPTED_HARNESS + textwrap.dedent(r"""
    const _passThrough = globalThis.fetch;
    globalThis.fetch = async (url, opts = {}) => {
      const u = String(url);
      if ((opts.method || 'GET') === 'GET' && u.includes('roledefinitions')
          && u.includes('$filter=Name')) {
        calls.push({ url: u, method: 'GET', body: null });
        const payload = { error: { message: { value: 'probe exploded' } } };
        return {
          ok: false, status: 500,
          headers: { get: () => null },
          json: async () => payload,
          text: async () => JSON.stringify(payload),
        };
      }
      return _passThrough(url, opts);
    };
""")


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_survey_failure_still_produces_a_structured_abort() -> None:
    """`_run_group_verify_deploy` already asserts a `__RESULT__` line was
    printed; if the probe failure above escaped as an unhandled rejection,
    that assertion is what would catch it, not the abort-key check below.
    """
    summary, calls, output = _run_group_verify_deploy(
        _deploy_js(), _SURVEY_FAILURE_HARNESS,
    )
    assert not _security_writes(calls), output[-2000:]
    errors = _security_errors(summary)
    assert errors, summary
    message = str(errors[0]["error"])
    assert "Schema Manager" in message, message
    assert "500" in message, message
    assert summary.get("aborted") == "phase-0-security-errors", summary


def _group_create_writes(calls: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Every POST that creates a new site group."""
    return [
        c for c in calls
        if c["method"] == "POST" and c["body"]
        and c["url"].endswith("/web/sitegroups")
    ]


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_clean_run_still_writes_both_the_level_and_the_group() -> None:
    """No refusals, no survey failures: the restructure must not turn a
    previously clean run into one that skips writes it used to make.

    Plain `_ADOPTED_HARNESS` carries no entry in `KNOWN_GROUP_NAMES`, so
    'List Maintainer' takes the create path here rather than the adopt
    path other tests in this file exercise via `_group_gate_deploy`.
    """
    summary, calls, _ = _run_group_verify_deploy(_deploy_js(), _ADOPTED_HARNESS)
    assert not _security_errors(summary), summary
    assert _role_def_merge_writes(calls, "Schema Manager"), (
        "a clean run stopped reconciling the permission level"
    )
    assert _group_create_writes(calls), (
        "a clean run stopped creating the site group"
    )


# Task 5 (#32): the decision table. Every object BOTH survey loops decided
# to create or adopt must be named before either loop's apply step writes
# anything. `_ADOPTED_HARNESS` carries no entry in `KNOWN_GROUP_NAMES`, so
# the group takes the create path and the level (matching this family's
# marker) takes the adopt path in the same run, exercising both verbs.


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_the_decision_table_names_every_declared_object_before_any_write() -> None:
    """A refusal already logs its own ERROR line in the survey loop that
    found it; this table is what a CLEAN run additionally gets, printed
    before a single write.
    """
    summary, _calls, output = _run_group_verify_deploy(_deploy_js(), _ADOPTED_HARNESS)
    assert not _security_errors(summary), summary
    lines = output.splitlines()

    table_index = next(
        (i for i, ln in enumerate(lines)
         if "decisions" in ln.lower() and pn("security") in ln), None,
    )
    assert table_index is not None, f"no decision table was printed:\n{output[-2000:]}"

    table_block = "\n".join(lines[table_index:table_index + 5])
    assert "Schema Manager" in table_block, table_block
    assert "List Maintainer" in table_block, table_block

    first_write_log = next(
        i for i, ln in enumerate(lines)
        if "Creating" in ln or "reconciled" in ln
    )
    assert table_index < first_write_log, (
        f"the decision table printed after a write had already started:\n{output[-2000:]}"
    )


def _duplicate_group_case_variant(js: str) -> str:
    """Splice a second declared group into `SCHEMA.groups`, differing from
    the first only in case, so `surveyGroup` meets a name `decidedCreates`
    already holds. `sharepoint-mapping.yaml` declares one group ('List
    Maintainer'); the build itself refuses two case-variant declarations in
    one mapping (`DUPLICATE_GROUP_NAME`), so this bypasses that by editing
    the already-generated JSON rather than the mapping, modelling a bundle
    built before that rule existed.
    """
    match = re.search(r'  "groups": (\[\n.*?\n  \]),\n  "indexed_columns"', js, re.DOTALL)
    assert match, "SCHEMA.groups block not found in generated deploy.js"
    groups = json.loads(match.group(1))
    variant = dict(groups[0])
    variant["name"] = "LIST MAINTAINER"
    return js.replace(match.group(1), json.dumps([*groups, variant], indent=2), 1)


@pytest.mark.skipif(NODE is None, reason="node is not installed")
def test_a_case_variant_group_declaration_is_refused_not_double_created() -> None:
    """`decidedCreates` used to feed only `isKnown`, which only decides
    whether to skip the by-name probe. For a name already decided 'create',
    the group does not exist yet, so the real probe still answers 404 and
    the survey returned a SECOND 'create' decision: applied, that queues two
    POSTs colliding on the one name SharePoint resolves them both to.

    The fix refuses the second declaration in the survey itself: one create
    decision reaches the table, the second is a refusal, and because a
    refusal blocks the whole phase's apply step, neither group is actually
    written. Mutation-tested: deleting the `hasName(decidedCreates, ...)`
    check in `surveyGroup` makes this test fail, printing two 'create site
    group' lines and no case-collision error.
    """
    js = _duplicate_group_case_variant(_deploy_js())
    summary, calls, output = _run_group_verify_deploy(js, _ADOPTED_HARNESS)

    create_lines = [ln for ln in output.splitlines() if "create site group" in ln]
    assert len(create_lines) == 1, (
        f"expected exactly one create decision for the two case-variant "
        f"declarations, got: {create_lines}"
    )
    errors = _security_errors(summary)
    assert len(errors) == 1, summary
    assert "group" in errors[0], errors[0]
    message = str(errors[0]["error"])
    assert "LIST MAINTAINER" in message, message
    assert "case" in message.lower(), message
    assert not _group_create_writes(calls), (
        "a refusal must block the apply step entirely, not just the refused object"
    )
    assert not _role_def_merge_writes(calls, "Schema Manager"), (
        "a refusal must block every object's apply, not only the colliding group's"
    )
    assert summary.get("aborted") == "phase-0-security-errors", summary


def test_no_reader_no_enrolment_code() -> None:
    """Opt-in: the code path must not exist unless asked for.

    Absence, asserted on the emitted text, is the one thing a text
    assertion states exactly. A guard that is present but unreachable is
    the failure mode this file avoids elsewhere; a call site that is not
    emitted at all cannot run.
    """
    js = _deploy_js()
    assert "ensureuser" not in js
    assert f"Starting Phase {pn('reader_enrolment')}" not in js
    # And the same mapping, WITH a reader, does emit it. Otherwise the two
    # assertions above would also hold for a template that never works.
    assert "ensureuser" in _reader_deploy_js()


def test_the_deploy_confirms_the_editor_still_refuses_the_guard() -> None:
    """The emitted script must ask the tenant, once, rather than assume.

    Measured 2026-08-17 (view-edit-page-probe.js): a view is protected when
    its edit page returns 200 from the endpoint asked for, carries a
    sentinel, and does not carry the editor's control names. The guard is
    identical across views and each view's stored ViewQuery is already
    verified, so one page answers for all of them.
    """
    script = _deploy_js()
    assert "ViewEdit.aspx" in script
    assert 'name=\"FieldPicker1\"' in script
    # A sentinel gates the absence check. C6 measured a request for a view
    # that does not exist answering 200 with no editor controls, so absence
    # alone would call a page that is not a view protected. Pinned as the
    # declaration rather than as a substring: `ctl00` and `ViewEdit` are on
    # that page too and are named in the comment beside it, so a bare
    # containment test would pass on either of them.
    assert "const EDITOR_PAGE_SENTINEL = 'ViewFilter';" in script
    # English display text must not be the predicate: it reads correctly on
    # an English tenant and silently wrong on any other.
    assert "complex filter" not in script


def test_an_unreadable_settings_page_warns_and_a_readable_one_can_fail() -> None:
    """Unverifiable and unprotected must not collapse into one outcome.

    A redirect, a missing sentinel or a throw means the check could not
    answer, which is not evidence and must not abort a deploy whose views
    are otherwise verified. The editor's control being PRESENT is a
    determination that the protection did not take, so that one fails.
    """
    script = _deploy_js()
    start = script.index("async function confirmEditorRefusesTheGuard")
    block = script[start:script.index("await confirmEditorRefusesTheGuard();")]

    # The failure is reachable ONLY under the control-present test, and the
    # push is a statement rather than a guarded expression: `void 0 && push`
    # would satisfy a containment test while never running.
    condition = block.index("if (present.length > 0)")
    pushes = [
        line.strip() for line in block.splitlines() if "summary.errors.push" in line
    ]
    assert pushes == ["summary.errors.push({"]
    assert block.index("summary.errors.push") > condition
    assert block.index("log('ERROR'") > condition

    # Every path that could not answer returns without recording a failure.
    # A check that cannot see the page is not evidence the page is wrong.
    warns = [i for i, line in enumerate(block.splitlines()) if "log('WARN'" in line]
    assert len(warns) == 3
    lines = block.splitlines()
    for index in warns:
        following = "\n".join(lines[index:index + 6])
        assert "return;" in following
        assert "summary.errors.push" not in following
