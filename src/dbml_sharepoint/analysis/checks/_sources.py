# src/dbml_sharepoint/analysis/checks/_sources.py
"""Retention policies and enum sources."""

from dbml_sharepoint.analysis.checks._context import ValidationContext
from dbml_sharepoint.analysis.findings import FindingCode, Location, Section
from dbml_sharepoint.analysis.validator import Finding


def check(vc: ValidationContext) -> list[Finding]:
    bundle = vc.bundle
    table_names = vc.table_names
    enum_by_name = vc.enum_by_name
    findings: list[Finding] = []
    # Retention cross-checks only apply when retention config is actually
    # present. mapping_loader ties retention_policies and
    # retention_list_defaults together (both loaded from the same
    # retention-policies.yaml; both empty when retention_policies_source is
    # unset), but the checks below key off retention_policies specifically
    # so a bundle with list_defaults but no defined policies does not
    # spuriously error on every 'not in policies' comparison.
    if bundle.retention_policies:
        # Every retention list_default key must be a known entity.
        # Keys carry the APP_ prefix; DBML table names are unprefixed.
        prefix = bundle.mapping.prefix
        for entity_name in bundle.retention_list_defaults:
            bare = (
                entity_name.removeprefix(prefix)
                if prefix and entity_name.startswith(prefix)
                else entity_name
            )
            if bare not in table_names:
                findings.append(Finding(
                    FindingCode.UNKNOWN_ENTITY,
                    f"retention list_defaults references unknown entity: {entity_name}",
                    # Prose, not a dotted path: these two messages were written
                    # before there was a convention, so `location.path`
                    # ("retention[X]") is not the prefix. The location is the
                    # structured truth; the message is what has to change, and
                    # it cannot until nothing matches on the prose.
                    location=Location(Section.RETENTION, entity=entity_name),
                ))

        # Every list_defaults policy must be a known retention policy.
        for entity_name, policy in bundle.retention_list_defaults.items():
            if policy not in bundle.retention_policies:
                findings.append(Finding(
                    FindingCode.UNKNOWN_RETENTION_POLICY,
                    f"retention list_defaults[{entity_name}] = {policy} not in policies",
                    location=Location(Section.RETENTION, entity=entity_name),
                ))

    # For every configured enum_sources entry, cross-check its choices
    # against the matching DBML enum. A configured enum name with no
    # matching DBML enum is only a warning — the schema simply hasn't
    # defined that enum yet, which isn't by itself wrong. A DBML enum whose
    # members differ from the configured choices IS an error: SP would
    # render different choices from what the loader feeds to other enum
    # consumers.
    for enum_name, choices in bundle.enum_choices.items():
        dbml_enum = enum_by_name.get(enum_name)
        # Neither message's prefix is `location.path` either: the first quotes
        # the enum name inside the brackets (`enum_sources['Status']`) and the
        # second opens with the DBML enum rather than with the section. The
        # location says where the finding is; the messages say it differently
        # and cannot be reworded yet.
        at_enum = Location(Section.ENUM_SOURCES, entity=enum_name)
        if dbml_enum is None:
            findings.append(Finding(
                FindingCode.ENUM_SOURCE_HAS_NO_DBML_ENUM,
                f"enum_sources[{enum_name!r}] has no matching DBML enum "
                f"{enum_name!r} in the schema.",
                location=at_enum,
            ))
        elif dbml_enum.members != choices:
            findings.append(Finding(
                FindingCode.ENUM_MEMBERS_DIFFER,
                f"DBML enum {enum_name!r} members differ from configured "
                f"enum_sources[{enum_name!r}] -- "
                f"DBML: {dbml_enum.members!r}; YAML: {choices!r}",
                location=at_enum,
            ))

    return findings
