---
title: Findings
sidebar_position: 4
---

# Finding catalogue

Every rule this build can report. The `code` is the stable identity;
the message is prose and may be reworded.

| Code | Severity | What it means |
|---|---|---|
| `auto_increment_pk_must_be_id` | error | An auto-increment primary key is named something other than `Id`. |
| `column_name_too_long` | error | A column's internal name exceeds SharePoint's length limit. |
| `cross_site_expansion_unhandled` | error | A cross-site reference column needs an extension that expands it; the active one deferred. |
| `default_not_an_enum_member` | error | A column's default is not a member of the enum it is typed as. |
| `duplicate_column_name` | error | A table declares the same column name twice. |
| `duplicate_enum_name` | error | Two enums share a name. |
| `duplicate_table_name` | error | Two tables share a name. |
| `empty_enum` | warning | An enum declares no members. |
| `illegal_column_name_character` | error | A column name contains a character SharePoint rejects. |
| `legacy_choice_type` | error | A column uses the legacy `choice` type instead of a named DBML enum. |
| `orphan_enum` | warning | An enum is defined but no column references it. |
| `reserved_column_name` | error | A column uses a name SharePoint reserves. |
| `unique_unsupported_for_type` | error | `[unique]` is declared on a type SharePoint cannot enforce it for. |
| `unique_without_not_null` | warning | `[unique]` without `not null`, so uniqueness is enforced only on populated values. |
| `unknown_column_type` | error | A column's DBML type is not one the typemap knows. |
| `unknown_ref_target` | error | A `ref` points at a table the schema does not define. |
