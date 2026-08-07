# src/dbml_sharepoint/analysis/checks/_permissions.py
"""Permission levels, groups, and per-list policies."""

from dbml_sharepoint.analysis.checks._context import ValidationContext
from dbml_sharepoint.analysis.findings import FindingCode, Location, Section
from dbml_sharepoint.analysis.permissions import BASE_PERMISSIONS, BUILT_IN_LEVELS
from dbml_sharepoint.analysis.validator import (
    _ASSOCIATED_GROUP_ALIASES,
    _BUILTIN_SP_GROUPS,
    Finding,
)
from dbml_sharepoint.model.mapping_loader import ListPermissionPolicy

# These messages spell the level or group name with `!r`, so the quotes are
# inside the bracket -- `permission_levels['Reader']`. A Location holding
# `Reader` would render `permission_levels[Reader]` and stop being the
# message's prefix, and a Location holding `'Reader'` would be storing a
# repr in a data field. The section-level Location is the honest one until
# the messages themselves drop the `!r`.
_LEVELS = Location(Section.PERMISSION_LEVELS)
_GROUPS = Location(Section.GROUPS)
#: `list_permissions` is not one of the mapping's per-entity sections; its
#: paths are `list_permissions.default...` and `list_permissions.overrides`.
_DEFAULT_POLICY = Location(Section.LIST_PERMISSIONS, sub="default")
_OVERRIDES = Location(Section.LIST_PERMISSIONS, sub="overrides")


def check(vc: ValidationContext) -> list[Finding]:
    bundle = vc.bundle
    table_names = vc.table_names
    findings: list[Finding] = []
    # === Permissions cross-checks (R4) ===
    if bundle.mapping.permissions is not None:
        perms = bundle.mapping.permissions

        # list_permissions.default.site_role, when declared, must be a known
        # site role — a typo would silently scope the default to nothing. The
        # valid roles are data-driven: those declared on the mapping's
        # entities, plus "default" (no hardcoded any labels you choose.
        scope = perms.default_policy_site_role
        known_roles = {e.site_role for e in bundle.mapping.entities.values()} | {"default"}
        if scope is not None and scope not in known_roles:
            findings.append(Finding(
                FindingCode.UNKNOWN_SITE_ROLE,
                f"list_permissions.default.site_role: unknown site role "
                f"{scope!r} (mapping declares: {', '.join(sorted(known_roles))}).",
                location=Location(
                    Section.LIST_PERMISSIONS, sub="default.site_role",
                ),
            ))

        # permission_levels[*].name must be unique — CASE-INSENSITIVELY,
        # because that is how SharePoint resolves and de-duplicates them.
        # Two declarations differing only in case are one object on the
        # site: the second create fails on a name collision, mid-deploy,
        # after the first has already been made.
        seen_level_names: dict[str, str] = {}
        for lvl in perms.levels:
            key = lvl.name.casefold()
            if key in seen_level_names:
                findings.append(Finding(
                    FindingCode.DUPLICATE_PERMISSION_LEVEL_NAME,
                    f"permission_levels: duplicate name {lvl.name!r}"
                    + (f" (differs from {seen_level_names[key]!r} only in case; "
                       f"SharePoint treats them as one)"
                       if seen_level_names[key] != lvl.name else ""),
                    location=_LEVELS,
                ))
            seen_level_names.setdefault(key, lvl.name)

        # permission_levels[*].base_permissions must be known bits.
        for lvl in perms.levels:
            for bit in lvl.base_permissions:
                if bit not in BASE_PERMISSIONS:
                    findings.append(Finding(
                        FindingCode.UNKNOWN_BASE_PERMISSION,
                        f"permission_levels[{lvl.name!r}]: unknown base permission {bit!r}",
                        location=_LEVELS,
                    ))

        # groups[*].name must be unique — case-insensitively, for the same
        # reason as the levels above.
        seen_group_names: dict[str, str] = {}
        custom_group_names = {g.name for g in perms.groups}
        for grp in perms.groups:
            key = grp.name.casefold()
            if key in seen_group_names:
                findings.append(Finding(
                    FindingCode.DUPLICATE_GROUP_NAME,
                    f"groups: duplicate name {grp.name!r}"
                    + (f" (differs from {seen_group_names[key]!r} only in case; "
                       f"SharePoint treats them as one)"
                       if seen_group_names[key] != grp.name else ""),
                    location=_GROUPS,
                ))
            seen_group_names.setdefault(key, grp.name)

        # groups[*].owner_group must be a built-in SP group or a declared custom group.
        for grp in perms.groups:
            owner_ok = (
                grp.owner_group in _BUILTIN_SP_GROUPS
                or grp.owner_group in custom_group_names
            )
            if not owner_ok:
                findings.append(Finding(
                    FindingCode.UNKNOWN_OWNER_GROUP,
                    f"groups[{grp.name!r}]: owner_group {grp.owner_group!r} is not a "
                    f"built-in SP group or a declared custom group",
                    location=_GROUPS,
                ))

        # Collect all valid level names (built-in + declared custom).
        all_level_names = BUILT_IN_LEVELS | set(seen_level_names.values())
        # Collect all valid group names (declared custom + built-in SP groups).
        all_group_names = custom_group_names | _BUILTIN_SP_GROUPS

        def _check_policy_assignments(
            policy: ListPermissionPolicy, ctx: str, at: Location,
        ) -> None:
            """`ctx` renders the message prefix and `at` locates the finding.

            They are not the same string for an override: the message spells
            the entity with `!r` inside the bracket, which a Location cannot
            reproduce without storing a repr. `at` is the enclosing section
            instead, which is still a prefix of every message here.
            """
            for i, assignment in enumerate(policy.assignments):
                lvl = assignment.level
                if lvl not in all_level_names:
                    findings.append(Finding(
                        FindingCode.UNKNOWN_PERMISSION_LEVEL,
                        f"{ctx}.assignments[{i}]: level {lvl!r} is not a built-in "
                        f"or declared custom permission level",
                        location=at,
                    ))
                principal = assignment.principal
                if principal.kind != "group":
                    continue
                suggested_kind = _ASSOCIATED_GROUP_ALIASES.get(
                    (principal.name or "").casefold(),
                )
                if suggested_kind is not None:
                    # Phase 4.2 resolves kind=group via sitegroups/getbyname(name).
                    # The built-in aliases never exist under these literal names
                    # on real sites (the associated groups are named
                    # '<SiteTitle> Owners' etc.), so the assignment would fail
                    # at deploy time.
                    findings.append(Finding(
                        FindingCode.UNRESOLVABLE_ASSOCIATED_GROUP_ALIAS,
                        f"{ctx}.assignments[{i}]: principal group "
                        f"{principal.name!r} is a built-in associated-group "
                        f"alias that cannot be resolved by name at deploy time "
                        f"(real sites name it '<SiteTitle> ...'). Use principal "
                        f"kind {suggested_kind!r} instead.",
                        location=at,
                    ))
                elif principal.name not in all_group_names:
                    findings.append(Finding(
                        FindingCode.UNKNOWN_PRINCIPAL_GROUP,
                        f"{ctx}.assignments[{i}]: principal group {principal.name!r} "
                        f"is not a built-in or declared custom group",
                        location=at,
                    ))

        if perms.default_policy is not None:
            _check_policy_assignments(
                perms.default_policy, "list_permissions.default", _DEFAULT_POLICY,
            )

        # list_permissions.overrides keys must be unprefixed DBML table names.
        for entity_name, override_policy in perms.overrides.items():
            if entity_name not in table_names:
                # Distinct from UNKNOWN_ENTITY, which is about a name the
                # MAPPING does not declare. This one is about a name the DBML
                # does not declare, and the fix is usually the prefix.
                findings.append(Finding(
                    FindingCode.UNKNOWN_TABLE,
                    f"list_permissions.overrides: key {entity_name!r} is not a "
                    "known DBML table name (use unprefixed name, "
                    "e.g. 'Ticket' not 'APP_Ticket')",
                    location=_OVERRIDES,
                ))
            ctx_key = f"list_permissions.overrides[{entity_name!r}]"
            _check_policy_assignments(override_policy, ctx_key, _OVERRIDES)

    return findings
