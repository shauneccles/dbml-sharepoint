# src/dbml_sharepoint/analysis/findings.py
"""What a finding IS, separate from what produces one.

`checks/*` needs the vocabulary without importing the orchestrator, the same
layering rule that already forbids a generator importing from `checks/`.

The `code` is the identity. Everything keys off it: tests, the docs catalogue,
and `--explain`. The `message` is prose for a human and is free to be reworded
in any commit -- before this module existed, 294 test assertions matched
substrings of it, so it could not be.
"""

from dataclasses import dataclass
from enum import StrEnum
from typing import Literal

type Severity = Literal["error", "warning"]


class Section(StrEnum):
    """The mapping section a finding is about.

    These eighteen names were already being spelled into message prefixes by
    hand at 175 sites; this makes the set closed and the spelling checked.
    """

    CALCULATED_FORMULAS = "calculated_formulas"
    COLUMN_FORMATTING = "column_formatting"
    COLUMN_VALIDATION = "column_validation"
    CONDITIONS = "conditions"
    DEMO_ITEMS = "demo_items"
    DISPLAY_NAMES = "display_names"
    ENTITIES = "entities"
    ENUM_SOURCES = "enum_sources"
    FIELD_SETS = "field_sets"
    FORM_FORMATTING = "form_formatting"
    FORM_VISIBILITY = "form_visibility"
    GROUPS = "groups"
    LIST_VALIDATION = "list_validation"
    PERMISSION_LEVELS = "permission_levels"
    POLYMORPHIC_PATTERNS = "polymorphic_patterns"
    RETIRED_COLUMNS = "retired_columns"
    SCHEMA = "schema"
    VIEWS = "views"
    WATCHED_LISTS = "watched_lists"


@dataclass(frozen=True, slots=True)
class Location:
    """Where a finding is, as data rather than as a rendered prefix."""

    section: Section
    entity: str | None = None
    column: str | None = None
    view: str | None = None
    sub: str | None = None

    @property
    def path(self) -> str:
        """The dotted path these messages have always rendered by hand."""
        head = str(self.section)
        if self.entity is not None:
            head = f"{head}[{self.entity}]"
        tail = [p for p in (self.view, self.column, self.sub) if p is not None]
        return ".".join([head, *tail])


class FindingCode(StrEnum):
    """One member per rule. The catalogue of everything this tool can say.

    Adding a rule means adding a member here and a row in
    `website/docs/reference/findings.md`; `test_every_code_is_documented`
    enforces the pair.
    """

    # TEMPORARY. Scaffolding for the migration: every one of the 174
    # construction sites got this in one pass so the tree compiles while the
    # real codes are named a check module at a time. Task 3 of the
    # typed-boundaries plan deletes this member, and
    # `test_no_finding_is_unclassified` is how we will know it is gone.
    UNCLASSIFIED = "unclassified"

    UNKNOWN_ENTITY = "unknown_entity"

    # --- schema-only rules, from validator.validate() ---
    AUTO_INCREMENT_PK_MUST_BE_ID = "auto_increment_pk_must_be_id"
    COLUMN_NAME_TOO_LONG = "column_name_too_long"
    CROSS_SITE_EXPANSION_UNHANDLED = "cross_site_expansion_unhandled"
    DEFAULT_NOT_AN_ENUM_MEMBER = "default_not_an_enum_member"
    DUPLICATE_COLUMN_NAME = "duplicate_column_name"
    DUPLICATE_ENUM_NAME = "duplicate_enum_name"
    DUPLICATE_TABLE_NAME = "duplicate_table_name"
    EMPTY_ENUM = "empty_enum"
    ILLEGAL_COLUMN_NAME_CHARACTER = "illegal_column_name_character"
    LEGACY_CHOICE_TYPE = "legacy_choice_type"
    ORPHAN_ENUM = "orphan_enum"
    RESERVED_COLUMN_NAME = "reserved_column_name"
    UNIQUE_UNSUPPORTED_FOR_TYPE = "unique_unsupported_for_type"
    UNIQUE_WITHOUT_NOT_NULL = "unique_without_not_null"
    UNKNOWN_COLUMN_TYPE = "unknown_column_type"
    UNKNOWN_REF_TARGET = "unknown_ref_target"



@dataclass(frozen=True, slots=True)
class Finding:
    """One thing the build has to say about the declaration it was given."""

    code: FindingCode
    severity: Severity
    message: str
    location: Location | None = None
