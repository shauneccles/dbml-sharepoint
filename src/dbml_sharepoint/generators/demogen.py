# src/dbml_sharepoint/generators/demogen.py
"""Render demo-data.js (declared demo/sample rows, emitted with --seed).

The plan is generation-time typed: each field carries a `kind` so the
script knows whether to write a literal, resolve the deploying operator
(person columns take `<Name>Id`), resolve a demo_ref to a created item's
Id (lookups also take `<Name>Id`), or compute a run-time date from a
`today+/-N` offset. Cadence-derived demo surfaces (Review due, overdue
formatting, Tolerance due) must land on whatever day the demo runs.
The '[DEMO] ' Title marker (validated mandatory) is the in-record notice
and the teardown contract.
"""

from typing import Any

from dbml_sharepoint.analysis.ordering import site_tables_in_order
from dbml_sharepoint.analysis.typemap import (
    MULTI_VALUE_METADATA_TYPE,
    TODAY_SENTINEL,
    element_type,
    is_hyperlink,
    is_multi_value,
    is_person,
)
from dbml_sharepoint.model.mapping_types import MappingBundle
from dbml_sharepoint.model.parser import Schema
from dbml_sharepoint.model.release import Release
from dbml_sharepoint.templating import script_env

# The sentinel has one home (analysis/typemap.py) because this module and
# the validator must accept exactly the same language: the validator gates
# what may be declared, this decides what is generated, and a value one
# accepts and the other does not passes the build with zero findings and
# emits the literal string "today" into a script.
_TODAY_OFFSET = TODAY_SENTINEL

# The Title marker is the in-record demo notice: visible in every view and
# form header, and the marker rollback.js trusts. (Per-row list-item
# comments were tried and withdrawn: the modern Comments() endpoint is
# undocumented surface and rejected the write live, 2026-07-24, while
# adding nothing the marker doesn't already show.)
DEMO_TITLE_PREFIX = "[DEMO] "

_DATE_TYPES = {"date", "datetime"}


def _field_plan(col_type: str | None, name: str, value: Any) -> dict[str, Any] | None:
    """The typed plan for one demo field, or None when the field is omitted."""
    # Keyed on `demo_ref` rather than on being a dict at all: a hyperlink
    # value is also a mapping, and a bare isinstance check claimed it as a
    # lookup reference and then raised KeyError.
    if isinstance(value, dict) and "demo_ref" in value:
        return {"name": name, "kind": "ref", "value": str(value["demo_ref"])}
    if is_person(col_type):
        return {"name": name, "kind": "me", "value": None}
    if is_hyperlink(col_type):
        # A SharePoint URL column is a RECORD over REST (SP.FieldUrlValue,
        # Url + Description), not a scalar. Writing a bare string is
        # rejected at create time. Without this kind, a hyperlink column
        # simply could not be seeded, which is why four templates in the
        # people theme shipped their EvidenceUrl and MinutesUrl blank.
        #
        # Authored as either "https://..." or {url: ..., description: ...};
        # a bare string takes the URL as its own description, which is what
        # SharePoint shows when an author leaves the description empty.
        if isinstance(value, dict):
            raw_url, description = value.get("url"), value.get("description")
        else:
            raw_url, description = value, None
        # Never str() a value that might be None: it yields "None", which is
        # a perfectly valid-looking string and becomes a link to nowhere.
        # The validator refuses this shape, so reaching here with a non-string
        # means the two readers have drifted. Fail rather than emit.
        if not isinstance(raw_url, str) or not raw_url.strip():
            raise ValueError(
                f"{name}: a hyperlink demo value needs a non-empty url, got {raw_url!r}",
            )
        url = raw_url
        return {
            "name": name,
            "kind": "url",
            "value": {"url": url, "description": str(description or url)},
        }
    # Through `element_type`, because `date[]` is not a key in this set and the
    # test read as though it covered the column.
    if element_type(col_type or "") in _DATE_TYPES and isinstance(value, str):
        m = _TODAY_OFFSET.match(value)
        if m:
            # The shared pattern captures sign and digits separately, so an
            # offset is rebuilt from both rather than read from one group.
            sign, digits = m.group(1), m.group(2)
            offset = int(digits) if digits else 0
            return {
                "name": name,
                "kind": "date_offset",
                "value": -offset if sign == "-" else offset,
            }
    # Placed immediately before `literal`, which is where a multi-value column
    # would otherwise land and emit the bare array nothing has sent. The kinds
    # above keep their place: each answers a shape arity does not change, and
    # `date[]` reaches the date grammar deliberately.
    if is_multi_value(col_type or ""):
        if not isinstance(value, list):
            # DEMO_MULTI_VALUE_NOT_A_LIST reports this first, so reaching here
            # means the validator and the planner have drifted. There is no
            # honest coercion: one member is not a measured collection.
            raise ValueError(
                f"{name}: a multi-value demo value must be a list of members, "
                f"got {value!r}.",
            )
        if not value:
            # An empty list OMITS the field. `multi-value-probe.js:586` seeded
            # its empty row behind `if (row.values.length)` and M4 measured
            # that column reading back `null` (2026-08-17), so omission is the
            # only route to an unset column that anything has sent.
            return None
        return {
            "name": name,
            "kind": "multi_value",
            "metadata_type": MULTI_VALUE_METADATA_TYPE,
            "results": list(value),
        }
    return {"name": name, "kind": "literal", "value": value}


def generate_demo_js(
    *,
    schema: Schema,
    bundle: MappingBundle,
    release: Release,
    site_url: str,
    site_role: str,
    source_dbml: str,
    generated_at: str,
) -> str:
    env = script_env()
    tables_by_name = {t.name: t for t in schema.tables}
    demo_plan: list[dict[str, Any]] = []
    for table_name in site_tables_in_order(schema, bundle.mapping.entities, site_role):
        table = tables_by_name[table_name]
        types_by_col = {c.name: c.type for c in table.columns}
        for item in bundle.mapping.demo_items.get(table_name, []):
            planned = (
                _field_plan(types_by_col.get(name), name, value)
                for name, value in item.values.items()
            )
            demo_plan.append({
                "list": bundle.mapping.prefix + table_name,
                "key": item.key,
                # A None plan is an omitted field, which is how an empty
                # multi-value column stays unset.
                "fields": [field for field in planned if field is not None],
            })

    template = env.get_template("demo.js.j2")
    return template.render(
        site_url=site_url,
        site_role=site_role,
        release=release,
        source_dbml=source_dbml,
        generated_at=generated_at,
        demo_plan=demo_plan,
        demo_title_prefix=DEMO_TITLE_PREFIX,
    )
