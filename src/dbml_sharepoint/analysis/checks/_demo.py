# src/dbml_sharepoint/analysis/checks/_demo.py
"""Demo rows seeded by ``--seed``."""

import datetime as dt
import re

from dbml_sharepoint.analysis.checks.context import ValidationContext
from dbml_sharepoint.analysis.findings import Finding, FindingCode, Location, Section
from dbml_sharepoint.analysis.rendered_columns import rendered_columns
from dbml_sharepoint.analysis.typemap import (
    DATE_TYPES,
    TODAY_SENTINEL,
    choice_enum_for,
    element_type,
    is_hyperlink,
    is_multi_value,
    is_person,
)

_DEMO_ISO_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


def check(vc: ValidationContext) -> list[Finding]:
    bundle = vc.bundle
    tables_by_name = vc.tables_by_name
    enum_by_name = vc.enum_by_name
    cross_site_by_entity = vc.cross_site_by_entity
    findings: list[Finding] = []
    # Demo rows (--seed): everything checkable at build time IS checked.
    # A demo paste failing live in front of an audience is exactly the
    # failure class this tool exists to prevent.
    demo_keys: dict[str, str] = {}
    for entity_name, demo_rows in bundle.mapping.demo_items.items():
        if entity_name not in tables_by_name or entity_name not in bundle.mapping.entities:
            findings.append(Finding(
                FindingCode.UNKNOWN_ENTITY, f"demo_items[{entity_name}]: unknown entity.",
                location=Location(Section.DEMO_ITEMS, entity=entity_name),
            ))
            continue
        # A document library's items ARE files. demo-data.js POSTs to
        # /items, and SharePoint refuses that on a library outright:
        # HTTP 500, "To add an item to a document library, use
        # SPFileCollection.Add()" (probed 2026-07-29,
        # test/manual/document-library-probe.js, L2).
        #
        # So the paste fails, loudly, in front of whoever was being shown
        # the demo. Refusing at build turns that into a failed build.
        # Seeding a library would mean uploading real files, which is a
        # different feature from writing list rows and is not one this tool
        # has.
        if bundle.mapping.entities[entity_name].kind == "DocumentLibrary":
            findings.append(Finding(
                FindingCode.DEMO_ROWS_ON_DOCUMENT_LIBRARY,
                f"demo_items[{entity_name}]: {entity_name} is a DocumentLibrary, and a "
                f"library's items are files. Seeding posts to /items, which SharePoint "
                f"refuses outright -- HTTP 500, \"To add an item to a document library, "
                f"use SPFileCollection.Add()\" -- so the paste fails in front of whoever "
                f"was being shown the demo. Seed the register list that accompanies the "
                f"library, and upload sample documents by hand.",
                location=Location(Section.DEMO_ITEMS, entity=entity_name),
            ))
            continue
        for row in demo_rows:
            if row.key in demo_keys:
                findings.append(Finding(
                    FindingCode.DUPLICATE_DEMO_KEY,
                    f"demo_items[{entity_name}].{row.key}: duplicate demo "
                    f"key (also declared under {demo_keys[row.key]}).",
                    location=Location(
                        Section.DEMO_ITEMS, entity=entity_name, sub=row.key,
                    ),
                ))
            else:
                demo_keys[row.key] = entity_name
    for entity_name, demo_rows in bundle.mapping.demo_items.items():
        demo_table = tables_by_name.get(entity_name)
        if demo_table is None or entity_name not in bundle.mapping.entities:
            continue
        xcols = cross_site_by_entity.get(entity_name, set())
        demo_writable = rendered_columns(demo_table, xcols) | {"Title"}
        demo_types = {c.name: c.type for c in demo_table.columns}
        columns = {c.name: c for c in demo_table.columns}
        row_positions = {row.key: position for position, row in enumerate(demo_rows)}
        for position, row in enumerate(demo_rows):
            ctx = f"demo_items[{entity_name}].{row.key}"
            at = Location(Section.DEMO_ITEMS, entity=entity_name, sub=row.key)
            demo_title = row.values.get("Title")
            if not isinstance(demo_title, str) or not demo_title.startswith("[DEMO] "):
                findings.append(Finding(
                    FindingCode.DEMO_TITLE_MISSING_MARKER,
                    f"{ctx}: Title must start with '[DEMO] ' -- the marker "
                    f"the teardown trusts to tell demo rows from real "
                    f"records.",
                    location=at,
                ))
            for col_name, value in row.values.items():
                col_type = demo_types.get(col_name)
                if col_name not in demo_writable or col_name == "Id":
                    findings.append(Finding(
                        FindingCode.DEMO_COLUMN_NOT_WRITABLE,
                        f"{ctx}: values references {col_name!r}, which is "
                        f"not a writable column of {entity_name}.",
                        location=at,
                    ))
                    continue
                if isinstance(col_type, str) and col_type.startswith("calculated"):
                    findings.append(Finding(
                        FindingCode.DEMO_VALUE_ON_CALCULATED_COLUMN,
                        f"{ctx}: {col_name} is a calculated column; demo "
                        f"rows cannot write it (set its inputs instead).",
                        location=at,
                    ))
                    continue
                # ARITY IS ANSWERED FIRST, because every rule below reads one
                # scalar and a rule handed a list reports the wrong fault.
                if is_multi_value(col_type or ""):
                    if not isinstance(value, list):
                        findings.append(Finding(
                            FindingCode.DEMO_MULTI_VALUE_NOT_A_LIST,
                            f"{ctx}: {col_name} is a multi-value column "
                            f"({col_type}), so its demo value must be a list of "
                            f"members, got {value!r}. An empty list is accepted "
                            f"and leaves the column unset.",
                            location=at,
                        ))
                        continue
                    # Compared by list rather than by set: a demo value can
                    # hold an unhashable member, and these lists are tiny.
                    seen: list[object] = []
                    repeated: list[object] = []
                    for member in value:
                        if member not in seen:
                            seen.append(member)
                        elif member not in repeated:
                            repeated.append(member)
                    if repeated:
                        findings.append(Finding(
                            FindingCode.DEMO_MULTI_VALUE_DUPLICATE_MEMBER,
                            f"{ctx}: {col_name} repeats "
                            f"{', '.join(repr(m) for m in repeated)}. The write "
                            f"shape measured as M3 on 2026-08-17 is a collection "
                            f"of choices and nothing "
                            f"has measured what a repeat reads back as, so declare "
                            f"each member once.",
                            location=at,
                        ))
                # Hyperlinks are checked before the object-value rule below,
                # and in BOTH authored shapes:
                # a URL column takes a bare address or a {url, description}
                # record. Gating this on `isinstance(value, dict)` left the
                # scalar form unchecked, and the generator DOES refuse a
                # non-string there, so an invalid mapping surfaced as a build
                # traceback rather than a finding. A validator must refuse
                # everything its generator refuses, and refuse it first.
                if is_hyperlink(col_type):
                    if isinstance(value, dict):
                        unknown = set(value) - {"url", "description"}
                        if unknown or "url" not in value:
                            findings.append(Finding(
                                FindingCode.DEMO_HYPERLINK_OBJECT_INVALID,
                                f"{ctx}: {col_name} is a hyperlink; an object value "
                                f"must be {{url: <address>, description: <label>}} "
                                f"with 'description' optional. Got keys "
                                f"{sorted(value)}.",
                                location=at,
                            ))
                            continue
                        address = value["url"]
                    else:
                        address = value
                    # Checked as a STRING, not stringified: str(None) is
                    # "None", which is non-empty, so a coerced emptiness test
                    # passes a null through to become a link pointing at the
                    # word None.
                    if not (isinstance(address, str) and address.strip()):
                        findings.append(Finding(
                            FindingCode.DEMO_HYPERLINK_ADDRESS_INVALID,
                            f"{ctx}: {col_name} is a hyperlink; its address must be "
                            f"a non-empty string, got {address!r}.",
                            location=at,
                        ))
                    continue
                if isinstance(value, dict):
                    if set(value) != {"demo_ref"}:
                        findings.append(Finding(
                            FindingCode.DEMO_OBJECT_VALUE_INVALID,
                            f"{ctx}: {col_name} object value must be exactly "
                            f"{{demo_ref: <key>}}.",
                            location=at,
                        ))
                    elif value["demo_ref"] not in demo_keys:
                        findings.append(Finding(
                            FindingCode.DEMO_REF_UNKNOWN_KEY,
                            f"{ctx}: {col_name} demo_ref "
                            f"{value['demo_ref']!r} is not a declared demo "
                            f"key.",
                            location=at,
                        ))
                    else:
                        column = columns.get(col_name)
                        target_entity = demo_keys[value["demo_ref"]]
                        if column is None or column.ref is None:
                            findings.append(Finding(
                                FindingCode.DEMO_REF_ON_NON_LOOKUP,
                                f"{ctx}: {col_name} uses demo_ref but is not a lookup column.",
                                location=at,
                            ))
                        elif column.ref.target_table != target_entity:
                            findings.append(Finding(
                                FindingCode.DEMO_REF_TARGET_MISMATCH,
                                f"{ctx}: {col_name} targets {column.ref.target_table}, but "
                                f"demo_ref {value['demo_ref']!r} belongs to {target_entity}.",
                                location=at,
                            ))
                        elif (
                            target_entity == entity_name
                            and row_positions.get(value["demo_ref"], position) >= position
                        ):
                            findings.append(Finding(
                                FindingCode.DEMO_REF_FORWARD_REFERENCE,
                                f"{ctx}: {col_name} demo_ref {value['demo_ref']!r} must be "
                                f"declared before the row that uses it.",
                                location=at,
                            ))
                    continue
                if is_person(col_type):
                    if value != "@me":
                        findings.append(Finding(
                            FindingCode.DEMO_PERSON_VALUE_UNSUPPORTED,
                            f"{ctx}: person column {col_name} accepts only "
                            f"\"@me\" (the deploying operator).",
                            location=at,
                        ))
                    continue
                # The enum question first: `_resolve_column` resolves an enum
                # before a scalar, so an enum named `date` is a Choice column
                # there and has to be one here too.
                enum_name = choice_enum_for(col_type or "", enum_by_name)
                # Through `element_type`, because `date[]` is not a key in
                # DATE_TYPES and the rule read as though it covered the column.
                if enum_name is None and element_type(col_type or "") in DATE_TYPES:
                    valid_date = False
                    if isinstance(value, str) and TODAY_SENTINEL.match(value):
                        valid_date = True
                    elif isinstance(value, str) and _DEMO_ISO_DATE.match(value):
                        try:
                            dt.date.fromisoformat(value)
                            valid_date = True
                        except ValueError:
                            pass
                    if not valid_date:
                        findings.append(Finding(
                            FindingCode.DEMO_DATE_VALUE_INVALID,
                            f"{ctx}: date column {col_name} accepts "
                            f"'today+N'/'today-N' or a real ISO calendar date "
                            f"(got {value!r}).",
                            location=at,
                        ))
                    continue
                # The resolver answers with the name; this rule needs the
                # members behind it.
                if enum_name is None:
                    continue
                demo_enum = enum_by_name[enum_name]
                # Membership is the same rule per member on a multi-value
                # column, so it keeps the same code rather than gaining a twin.
                # The arity test stays here: on a SCALAR enum column a list is
                # simply not a member, which is what already fires today.
                declared = (
                    value
                    if is_multi_value(col_type or "") and isinstance(value, list)
                    else [value]
                )
                for member in declared:
                    if member not in demo_enum.members:
                        findings.append(Finding(
                            FindingCode.DEMO_ENUM_VALUE_UNKNOWN,
                            f"{ctx}: {col_name} value {member!r} is not a member "
                            f"of enum {col_type}.",
                            location=at,
                        ))

    return findings
