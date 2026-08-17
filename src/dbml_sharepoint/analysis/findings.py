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
    CROSS_SITE_REFERENCE_COLUMNS = "cross_site_reference_columns"
    DEMO_ITEMS = "demo_items"
    DISPLAY_NAMES = "display_names"
    ENTITIES = "entities"
    ENUM_SOURCES = "enum_sources"
    FIELD_SETS = "field_sets"
    FORM_FORMATTING = "form_formatting"
    FORM_VISIBILITY = "form_visibility"
    GROUPS = "groups"
    # Not a per-entity mapping section like the rest: its paths are
    # `list_permissions.default...` and `list_permissions.overrides[...]`.
    LIST_PERMISSIONS = "list_permissions"
    LIST_VALIDATION = "list_validation"
    PERMISSION_LEVELS = "permission_levels"
    POLYMORPHIC_PATTERNS = "polymorphic_patterns"
    # Not one of the eighteen message prefixes: retention lives in its own
    # `retention-policies.yaml`, and its two findings were written as prose
    # rather than as a dotted path. The section is real even where the message
    # never spelled it.
    RETENTION = "retention"
    RETIRED_COLUMNS = "retired_columns"
    SCHEMA = "schema"
    VERSIONING = "versioning"
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
    """One member per rule, each declared WITH its severity.

    The catalogue of everything this tool can say, and the severity is part
    of the declaration rather than something a caller supplies.

    That is deliberate and it is the whole point of the two-value form. A
    rule's severity is a property of the rule -- `unique without not_null` is
    a warning because uniqueness still half-works, and no call site gets to
    have an opinion about that. When `Finding(...)` took a severity argument,
    every one of the 155 construction sites was an opportunity to disagree
    with the published catalogue, and one already did: the suite raised
    `extension_reported` as a warning while `findings.md` documented it as an
    error, and nothing could see it.

    Now there is nowhere to put a second answer. `Finding.severity` reads
    `code.severity`, the docs page is generated from the same member, and a
    construction site cannot pass one at all.

    A rule that genuinely comes in two strengths is two codes -- see
    `EXTENSION_REPORTED` and `EXTENSION_WARNING` -- which is more honest
    anyway: the reader of a code should not have to ask which of two things
    it meant this time.
    """

    def __new__(cls, value: str, severity: Severity) -> "FindingCode":
        member = str.__new__(cls, value)
        member._value_ = value
        member._severity = severity
        return member

    _severity: Severity

    @property
    def severity(self) -> Severity:
        """How bad this rule is. Declared with the member, never passed in."""
        return self._severity

    # An extension's own rule, via DeploymentExtension.extra_validators. The
    # core cannot enumerate what a project-specific validator will object to,
    # so they share these codes and carry the detail in the message.
    #
    # Two of them, because severity is fixed per code and an extension's
    # validator legitimately needs both. This is the one rule in the
    # catalogue whose strength the core cannot know, so it is the one rule
    # that is split rather than decided.
    EXTENSION_REPORTED = "extension_reported", "error"
    EXTENSION_WARNING = "extension_warning", "warning"

    # Reachable from `views`, `field_sets`, `display_names` and `retention`:
    # one rule, four sections, and the section is in the location.

    # --- checks/_views.py: field sets -------------------------------------
    # A name in the mapping is not a rendered column of the entity. Reached
    # from a field set's members, a view's fields/sort/group_by, and a
    # display-name override -- the same mistake, told apart by the location.
    # --- checks/_views.py: declared views ---------------------------------
    # The list view LOOKUP threshold, one pair of codes for one rule: the
    # same `_join_finding` is reached from a declared view and from the
    # generated All Items view, so the section is what tells them apart.
    # --- checks/_views.py: hide_from_all_items ----------------------------
    # --- checks/_naming.py ------------------------------------------------
    # --- checks/_sources.py -----------------------------------------------
    # A mapping section names an entity that does not resolve to a table. One
    # rule reachable from many sections, so one code -- `location.section` is
    # what says which section asked. `entities:` itself is NOT this code: see
    # ENTITY_NOT_IN_SCHEMA, which is about the declaration rather than a
    # reference to it.
    # --- checks/_structure.py: entities, cross-site refs, indexes, calculated
    ALL_ITEMS_VIEW_DECLARED = "all_items_view_declared", "error"
    CALCULATED_COLUMN_HAS_NO_FORMULA = "calculated_column_has_no_formula", "error"
    CALCULATED_DISPLAY_COLUMN_UNINDEXABLE = "calculated_display_column_unindexable", "warning"
    CALCULATED_FORMULA_CYCLE = "calculated_formula_cycle", "error"
    CALCULATED_FORMULA_DEFERRED_LOOKUP = "calculated_formula_deferred_lookup", "error"
    CALCULATED_FORMULA_MISSING_EQUALS = "calculated_formula_missing_equals", "error"
    CALCULATED_FORMULA_SELF_REFERENCE = "calculated_formula_self_reference", "error"
    CALCULATED_FORMULA_TOO_LONG = "calculated_formula_too_long", "error"
    CALCULATED_FORMULA_UNKNOWN_COLUMN = "calculated_formula_unknown_column", "error"
    CALCULATED_FORMULA_UNSUPPORTED_OPERAND = "calculated_formula_unsupported_operand", "error"
    COLUMN_NOT_RENDERED = "column_not_rendered", "error"
    COMPOSITE_INDEX_UNSUPPORTED = "composite_index_unsupported", "error"
    CROSS_SITE_COLUMN_CANNOT_BE_UNIQUE = "cross_site_column_cannot_be_unique", "error"
    CROSS_SITE_COLUMN_HAS_NO_REF = "cross_site_column_has_no_ref", "error"
    CROSS_SITE_GENERATED_NAME_COLLIDES = "cross_site_generated_name_collides", "error"
    CROSS_SITE_GENERATED_NAME_TOO_LONG = "cross_site_generated_name_too_long", "error"
    CROSS_SITE_UNKNOWN_COLUMN = "cross_site_unknown_column", "error"
    DISPLAY_COLUMN_NOT_RENDERED = "display_column_not_rendered", "error"
    DISPLAY_COLUMN_TYPE_UNINDEXABLE = "display_column_type_unindexable", "error"
    DISPLAY_TITLE_TOO_LONG = "display_title_too_long", "error"
    DOCUMENT_LIBRARY_UNSUPPORTED = "document_library_unsupported", "error"
    DUPLICATE_DISPLAY_TITLE = "duplicate_display_title", "error"
    DISPLAY_TITLE_COLLIDES_WITH_REPORT_COLUMN = (
        "display_title_collides_with_report_column", "error",
    )
    DUPLICATE_INDEX_TARGET = "duplicate_index_target", "error"
    DUPLICATE_VIEW_TITLE = "duplicate_view_title", "error"
    DUPLICATE_VIEW_URL_SLUG = "duplicate_view_url_slug", "error"
    EMPTY_DISPLAY_TITLE = "empty_display_title", "error"
    EMPTY_PREVIOUS_TITLE = "empty_previous_title", "error"
    EMPTY_VIEW_URL_SLUG = "empty_view_url_slug", "error"
    ENTITY_HAS_NO_NOTE = "entity_has_no_note", "error"
    ENTITY_NOT_IN_SCHEMA = "entity_not_in_schema", "error"
    ENTITY_NOTE_WHITESPACE_UNMEASURED = (
        "entity_note_whitespace_unmeasured", "error"
    )
    ENTITY_NOTE_TOO_LONG_FOR_MARKER = (
        "entity_note_too_long_for_marker", "error"
    )
    ENUM_MEMBERS_DIFFER = "enum_members_differ", "error"
    ENUM_SOURCE_HAS_NO_DBML_ENUM = "enum_source_has_no_dbml_enum", "warning"
    FIELD_SET_EMPTY = "field_set_empty", "error"
    FIELD_SET_NAME_HAS_MARKER = "field_set_name_has_marker", "error"
    FIELD_SET_UNREFERENCED = "field_set_unreferenced", "warning"
    FORMATTER_FIELD_NOT_DISPLAYED = "formatter_field_not_displayed", "error"
    FORMATTER_FIELD_NOT_RENDERED = "formatter_field_not_rendered", "error"
    FORMULA_TARGET_NOT_CALCULATED = "formula_target_not_calculated", "error"
    HIDE_IS_UNNECESSARY = "hide_is_unnecessary", "warning"
    HIDE_OF_CROSS_SITE_REFERENCE = "hide_of_cross_site_reference", "error"
    HIDE_OF_NON_JOIN_BEARING_COLUMN = "hide_of_non_join_bearing_column", "error"
    HIDE_OF_UNRENDERED_COLUMN = "hide_of_unrendered_column", "error"
    HIDE_WITHOUT_ALL_ITEMS_VIEW = "hide_without_all_items_view", "error"
    INDEX_COLUMN_NOT_RENDERED = "index_column_not_rendered", "error"
    INDEX_COLUMN_TYPE_UNINDEXABLE = "index_column_type_unindexable", "error"
    INDEX_DUPLICATES_UNIQUE_COLUMN = "index_duplicates_unique_column", "error"
    INDEX_LIMIT_APPROACHING = "index_limit_approaching", "warning"
    INDEX_LIMIT_EXCEEDED = "index_limit_exceeded", "error"
    INDEX_ON_CALCULATED_COLUMN = "index_on_calculated_column", "error"
    INDEX_SETTINGS_UNSUPPORTED = "index_settings_unsupported", "error"
    JOIN_THRESHOLD_APPROACHED = "join_threshold_approached", "warning"
    JOIN_THRESHOLD_EXCEEDED = "join_threshold_exceeded", "error"
    LOOKUP_CROSSES_SITE_ROLE = "lookup_crosses_site_role", "error"
    LOOKUP_DISPLAY_COLUMN_UNKNOWN = "lookup_display_column_unknown", "error"
    LOOKUP_WOULD_RENDER_BLANK = "lookup_would_render_blank", "error"
    MULTIPLE_DEFAULT_VIEWS = "multiple_default_views", "error"
    # Both specialise a rule that already fires, and both exist because the
    # generic REMEDY is wrong here rather than because the generic diagnosis
    # is. An index is the answer to an unindexed filter and a different
    # column is the answer to an unindexable type; on a multi-value column
    # the first fails the build and the second is not the only option, since
    # the same enum without the brackets is indexable.
    MULTI_VALUE_FILTERED_VIEW_UNINDEXABLE = (
        "multi_value_filtered_view_unindexable", "warning"
    )
    MULTI_VALUE_INDEX_UNSUPPORTED = "multi_value_index_unsupported", "error"
    # One rule, three modules: a calculated formula (checks/_structure.py), a
    # validation formula and a show/hide formula (both analysis/conditions.py,
    # reached from column_validation and form_visibility). All three refuse a
    # multi-value operand, so all three keep one code and `location.section`
    # says which surface asked -- the same arrangement the `condition_` codes
    # already have. CAML is deliberately not among them: a view filter over a
    # multi-value column is measured to work.
    MULTI_VALUE_OPERAND_UNSUPPORTED = "multi_value_operand_unsupported", "error"
    POLYMORPHIC_COLUMN_NOT_RENDERED = "polymorphic_column_not_rendered", "error"
    PREVIOUS_TITLE_CLAIMED_TWICE = "previous_title_claimed_twice", "error"
    PREVIOUS_TITLE_IS_A_CURRENT_TITLE = "previous_title_is_a_current_title", "error"
    PREVIOUS_TITLE_IS_OWN_TITLE = "previous_title_is_own_title", "error"
    PREVIOUS_TITLE_IS_RESERVED = "previous_title_is_reserved", "error"
    REDUNDANT_DISPLAY_COLUMN_ACCEPTANCE = "redundant_display_column_acceptance", "warning"
    RETIRED_COLUMN_IN_FIELD_SET = "retired_column_in_field_set", "warning"
    ROW_LIMIT_OUT_OF_RANGE = "row_limit_out_of_range", "error"
    TOTAL_COLUMN_NOT_DISPLAYED = "total_column_not_displayed", "error"
    TOTAL_NEEDS_NUMERIC_COLUMN = "total_needs_numeric_column", "error"
    TOTAL_ON_LOOKUP_COLUMN = "total_on_lookup_column", "error"
    TOTAL_ON_NON_ARITHMETIC_COLUMN = "total_on_non_arithmetic_column", "error"
    UNINDEXED_FILTER_COLUMNS = "unindexed_filter_columns", "warning"
    UNKNOWN_ENTITY = "unknown_entity", "error"
    UNKNOWN_FIELD_SET_REFERENCE = "unknown_field_set_reference", "error"
    UNKNOWN_RETENTION_POLICY = "unknown_retention_policy", "error"
    UNMAPPED_SCHEMA_TABLE = "unmapped_schema_table", "error"
    UNSUPPORTED_BASE_TEMPLATE = "unsupported_base_template", "error"
    WATCHED_COLUMN_NOT_RENDERED = "watched_column_not_rendered", "error"
    # A view's CustomFormatter is stored in the view schema XML, so a bare `&`
    # in it opens an entity reference and makes SharePoint's own document
    # malformed. Measured 2026-08-11: the view MERGE returns HTTP 500,
    # System.Xml.XmlException, at the exact position of the `&`. It aborts the
    # deployment, and nothing in the build or `node --check` can see it.
    VIEW_FORMATTER_XML_METACHARACTER = (
        "view_formatter_xml_metacharacter", "error",
    )
    WIDTH_COLUMN_NOT_DISPLAYED = "width_column_not_displayed", "error"
    WIDTH_OUT_OF_RANGE = "width_out_of_range", "error"

    # --- checks/_formatting.py: column formatting, style specs, form
    # formatting, list validation
    UNDEPLOYABLE_COLUMN_DECLARATION = "undeployable_column_declaration", "error"
    FORMATTER_COLUMN_NOT_RENDERED = "formatter_column_not_rendered", "error"
    FORMATTER_MISSING_ELMTYPE = "formatter_missing_elmtype", "error"
    # One rule, three sections: a formatter naming a [$Field] the entity does
    # not render. `location.section` says whether it was a column formatter, a
    # form part or a view formatter.
    STYLE_REQUIRES_CALCULATED = "style_requires_calculated", "error"
    STYLE_CALCULATED_TYPE_MISMATCH = "style_calculated_type_mismatch", "error"
    STYLE_ON_BOOLEAN_MATCHES_NOTHING = "style_on_boolean_matches_nothing", "error"
    # Its sibling above, one measurement later. The two rules share a cause --
    # a quoted-string comparison against a value that is not a string -- and
    # differ in what the reader is left looking at, which is the only thing a
    # code has to say. The boolean case renders UNSTYLED; the multi-value case
    # was watched on a tenant on 2026-08-10 rendering a neutral fill on every
    # row, which is worse, because a gap reads as a gap and a fill reads as a
    # verdict. The name says what was seen: not an absence, a false neutral.
    MULTI_VALUE_STYLE_RENDERS_A_FALSE_NEUTRAL = (
        "multi_value_style_renders_a_false_neutral", "error"
    )
    STYLE_MAP_KEY_NOT_IN_ENUM = "style_map_key_not_in_enum", "error"
    COLOR_BY_MAP_KEY_NOT_IN_ENUM = "color_by_map_key_not_in_enum", "error"
    TREND_AGAINST_NOT_RENDERED = "trend_against_not_rendered", "error"
    OVERDUE_GUARD_FIELD_NOT_RENDERED = "overdue_guard_field_not_rendered", "error"
    FORM_PART_REFERENCES_CALCULATED_COLUMN = "form_part_references_calculated_column", "error"
    FORM_SECTION_FIELD_NOT_RENDERED = "form_section_field_not_rendered", "error"
    FORM_SECTION_ENTIRELY_HIDDEN = "form_section_entirely_hidden", "error"
    FORM_COLUMNS_IN_NO_SECTION = "form_columns_in_no_section", "warning"
    LIST_VALIDATION_MESSAGE_TOO_LONG = "list_validation_message_too_long", "error"
    LIST_VALIDATION_FORMULA_TOO_LONG = "list_validation_formula_too_long", "error"
    # The condition grammar rejected the expression. `conditions.py` has 28
    # distinct reasons behind this and hands them back as prose; splitting
    # them into codes is that module's classification, not this one's.
    INVALID_CONDITION = "invalid_condition", "error"

    # --- checks/_demo.py: rows seeded by `--seed`
    DEMO_ROWS_ON_DOCUMENT_LIBRARY = "demo_rows_on_document_library", "error"
    DUPLICATE_DEMO_KEY = "duplicate_demo_key", "error"
    DEMO_TITLE_MISSING_MARKER = "demo_title_missing_marker", "error"
    DEMO_COLUMN_NOT_WRITABLE = "demo_column_not_writable", "error"
    DEMO_VALUE_ON_CALCULATED_COLUMN = "demo_value_on_calculated_column", "error"
    DEMO_HYPERLINK_OBJECT_INVALID = "demo_hyperlink_object_invalid", "error"
    DEMO_HYPERLINK_ADDRESS_INVALID = "demo_hyperlink_address_invalid", "error"
    DEMO_OBJECT_VALUE_INVALID = "demo_object_value_invalid", "error"
    DEMO_REF_UNKNOWN_KEY = "demo_ref_unknown_key", "error"
    DEMO_REF_ON_NON_LOOKUP = "demo_ref_on_non_lookup", "error"
    DEMO_REF_TARGET_MISMATCH = "demo_ref_target_mismatch", "error"
    DEMO_REF_FORWARD_REFERENCE = "demo_ref_forward_reference", "error"
    DEMO_PERSON_VALUE_UNSUPPORTED = "demo_person_value_unsupported", "error"
    DEMO_DATE_VALUE_INVALID = "demo_date_value_invalid", "error"
    DEMO_ENUM_VALUE_UNKNOWN = "demo_enum_value_unknown", "error"
    DEMO_MULTI_VALUE_NOT_A_LIST = "demo_multi_value_not_a_list", "error"
    DEMO_MULTI_VALUE_DUPLICATE_MEMBER = "demo_multi_value_duplicate_member", "error"

    # --- retirement and the sections it folds into (checks/_retirement.py) --
    CALCULATED_FORMULA_REFERENCES_A_RETIRED_COLUMN = (
        "calculated_formula_references_a_retired_column", "error"
    )
    COLUMN_VALIDATION_ON_A_RETIRED_COLUMN = "column_validation_on_a_retired_column", "error"
    COLUMN_VALIDATION_REFERENCES_OTHER_COLUMNS = (
        "column_validation_references_other_columns", "error"
    )
    LIST_VALIDATION_REFERENCES_A_RETIRED_COLUMN = (
        "list_validation_references_a_retired_column", "error"
    )
    RETIRED_COLUMN_NOT_RENDERED = "retired_column_not_rendered", "error"
    RETIRED_COLUMN_REQUIRED_WITH_A_DEFAULT = "retired_column_required_with_a_default", "warning"
    RETIRED_COLUMN_STILL_INDEXED = "retired_column_still_indexed", "warning"
    RETIRED_DATE_NOT_ISO = "retired_date_not_iso", "error"
    RETIREMENT_STRIPPED_A_DECLARATION = "retirement_stripped_a_declaration", "warning"
    RETIREMENT_WITHOUT_DISPLAY_NAMES = "retirement_without_display_names", "warning"
    SUPERSEDED_BY_IS_ITSELF_RETIRED = "superseded_by_is_itself_retired", "error"
    SUPERSEDED_BY_NAMES_THE_RETIRED_COLUMN = "superseded_by_names_the_retired_column", "error"
    SUPERSEDED_BY_NOT_RENDERED = "superseded_by_not_rendered", "error"
    UNDEPLOYABLE_DECLARATION_COLUMN = "undeployable_declaration_column", "error"
    VALIDATION_FORMULA_TOO_LONG = "validation_formula_too_long", "error"
    VALIDATION_MESSAGE_TOO_LONG = "validation_message_too_long", "error"
    VIEW_EMPTIED_BY_RETIREMENT = "view_emptied_by_retirement", "warning"

    # --- permission levels, groups and policies (checks/_permissions.py) ----
    DUPLICATE_GROUP_NAME = "duplicate_group_name", "error"
    DUPLICATE_PERMISSION_LEVEL_NAME = "duplicate_permission_level_name", "error"
    ENTERPRISE_READER_GROUP_ENROLS_THE_OPERATOR = (
        "enterprise_reader_group_enrols_the_operator", "error")
    ENTERPRISE_READER_GROUP_MEMBERS_MAY_EDIT_MEMBERSHIP = (
        "enterprise_reader_group_members_may_edit_membership", "error")
    ENTERPRISE_READER_GROUP_NOT_GRANTED = (
        "enterprise_reader_group_not_granted", "error")
    ENTERPRISE_READER_GROUP_OVER_PRIVILEGED = (
        "enterprise_reader_group_over_privileged", "error")
    ENTERPRISE_READER_GROUP_REQUIRES_EMPTY = (
        "enterprise_reader_group_requires_empty", "error")
    GROUP_AUTO_ACCEPT_WITHOUT_REQUESTS = (
        "group_auto_accept_without_requests", "error")
    GROUP_DESCRIPTION_TOO_LONG = "group_description_too_long", "error"
    GROUP_DESCRIPTION_TOO_LONG_FOR_MARKER = (
        "group_description_too_long_for_marker", "error")
    MARKER_LONGER_THAN_THE_FIELD = (
        "marker_longer_than_the_field", "error")
    MARKER_FIELD_HAS_RESERVED_TEXT = (
        "marker_field_has_reserved_text", "error")
    MARKER_FAMILY_MISSING = (
        "marker_family_missing", "error")
    MULTIPLE_ENTERPRISE_READER_GROUPS = (
        "multiple_enterprise_reader_groups", "error")
    PERMISSION_LEVEL_DESCRIPTION_TOO_LONG = (
        "permission_level_description_too_long", "error")
    PERMISSION_LEVEL_DESCRIPTION_TOO_LONG_FOR_MARKER = (
        "permission_level_description_too_long_for_marker", "error")
    PERMISSION_LEVEL_NOT_DIRECTLY_ASSIGNABLE = (
        "permission_level_not_directly_assignable", "error")
    PERMISSION_LEVEL_REDEFINES_A_BUILTIN = (
        "permission_level_redefines_a_builtin", "error")
    UNKNOWN_BASE_PERMISSION = "unknown_base_permission", "error"
    UNKNOWN_OWNER_GROUP = "unknown_owner_group", "error"
    UNKNOWN_PERMISSION_LEVEL = "unknown_permission_level", "error"
    UNKNOWN_PRINCIPAL_GROUP = "unknown_principal_group", "error"
    UNKNOWN_SITE_ROLE = "unknown_site_role", "error"
    #: The DBML does not declare this table. Distinct from UNKNOWN_ENTITY,
    #: which is a name the MAPPING does not declare.
    UNKNOWN_TABLE = "unknown_table", "error"
    UNRESOLVABLE_ASSOCIATED_GROUP_ALIAS = "unresolvable_associated_group_alias", "error"

    # --- form visibility (analysis/forms.py) --------------------------------
    FORM_VISIBILITY_CONDITION_UNREACHABLE = "form_visibility_condition_unreachable", "error"
    FORM_VISIBILITY_ON_A_CALCULATED_COLUMN = "form_visibility_on_a_calculated_column", "error"
    REQUIRED_COLUMN_HIDDEN_FROM_THE_NEW_FORM = "required_column_hidden_from_the_new_form", "error"
    REQUIRED_COLUMN_MAY_BE_HIDDEN_AT_CREATION = (
        "required_column_may_be_hidden_at_creation", "warning"
    )

    # --- the shared condition grammar (analysis/conditions.py) --------------
    # The prefix names the SUBJECT, not a section: these are reachable from
    # views, form_visibility, column_validation and list_validation alike,
    # and the section is in the location.
    CONDITION_COLUMN_TYPE_UNKNOWN = "condition_column_type_unknown", "error"
    CONDITION_DATE_IS_AN_UNQUOTED_YAML_DATETIME = (
        "condition_date_is_an_unquoted_yaml_datetime", "error"
    )
    CONDITION_DATE_UNPARSEABLE = "condition_date_unparseable", "error"
    CONDITION_DATE_WEARS_WHITESPACE = "condition_date_wears_whitespace", "error"
    CONDITION_FIELD_NOT_RENDERED = "condition_field_not_rendered", "error"
    CONDITION_LOOKUP_UNSUPPORTED_BY_TARGET = "condition_lookup_unsupported_by_target", "error"
    CONDITION_ME_OPERATOR_MEANINGLESS = "condition_me_operator_meaningless", "error"
    CONDITION_ME_TAKES_NO_PROPERTY = "condition_me_takes_no_property", "error"
    CONDITION_ME_UNSUPPORTED_BY_TARGET = "condition_me_unsupported_by_target", "error"
    CONDITION_MEASURE_NOT_APPLICABLE = "condition_measure_not_applicable", "error"
    CONDITION_MEASURE_UNKNOWN = "condition_measure_unknown", "error"
    CONDITION_MEASURE_UNRENDERABLE = "condition_measure_unrenderable", "error"
    CONDITION_NEEDLE_EMPTY = "condition_needle_empty", "error"
    CONDITION_NEGATION_UNRENDERABLE = "condition_negation_unrenderable", "error"
    CONDITION_NEGATIVE_TEXT_OPERATOR_UNRENDERABLE = (
        "condition_negative_text_operator_unrenderable", "error"
    )
    CONDITION_NOW_ON_A_DATE_COLUMN = "condition_now_on_a_date_column", "error"
    CONDITION_NOW_UNSUPPORTED_BY_TARGET = "condition_now_unsupported_by_target", "error"
    CONDITION_OPERAND_TYPE_UNSUPPORTED = "condition_operand_type_unsupported", "error"
    CONDITION_OPERATOR_NOT_NEGATABLE = "condition_operator_not_negatable", "error"
    CONDITION_OPERATOR_UNKNOWN = "condition_operator_unknown", "error"
    CONDITION_OPERATOR_UNRENDERABLE = "condition_operator_unrenderable", "error"
    CONDITION_OPERATOR_UNVERIFIED = "condition_operator_unverified", "error"
    CONDITION_PROPERTY_NOT_APPLICABLE = "condition_property_not_applicable", "error"
    CONDITION_PROPERTY_REQUIRED = "condition_property_required", "error"
    CONDITION_PROPERTY_UNKNOWN = "condition_property_unknown", "error"
    CONDITION_PROPERTY_UNRENDERABLE = "condition_property_unrenderable", "error"
    CONDITION_SENTINEL_WITH_A_SUBSTRING_OPERATOR = (
        "condition_sentinel_with_a_substring_operator", "error"
    )
    CONDITION_SET_EMPTY = "condition_set_empty", "error"
    CONDITION_SUBSTRING_TEST_ON_A_NON_TEXT_COLUMN = (
        "condition_substring_test_on_a_non_text_column", "error"
    )
    CONDITION_TODAY_UNSUPPORTED_BY_TARGET = "condition_today_unsupported_by_target", "error"
    CONDITION_TOO_DEEP = "condition_too_deep", "error"
    CONDITION_TOO_MANY_LEAVES = "condition_too_many_leaves", "error"
    CONDITION_VALUE_HAS_A_CONTROL_CHARACTER = "condition_value_has_a_control_character", "error"
    CONDITION_VALUE_MISSING = "condition_value_missing", "error"
    CONDITION_VALUE_NOT_A_BOOLEAN = "condition_value_not_a_boolean", "error"
    CONDITION_VALUE_NOT_A_LIST = "condition_value_not_a_list", "error"
    CONDITION_VALUE_NOT_A_NUMBER = "condition_value_not_a_number", "error"
    CONDITION_VALUE_NOT_ALLOWED = "condition_value_not_allowed", "error"
    CONDITION_VALUE_NOT_FINITE = "condition_value_not_finite", "error"
    # Three rules about ARITY, raised from the same module and carrying the
    # `MULTI_VALUE_` prefix rather than `CONDITION_` for the reason stated
    # above: a code names its SUBJECT. A reader who meets one of these is
    # being told something about a multi-value column, not about the grammar,
    # and `explain multi_value_...` should list them together with the six
    # refusals that already exist.
    #
    # They come in a set of three because the defect has three faces, and all
    # three are the same trap: one authored word meaning two things depending
    # on a column's arity or a value's spelling.
    MULTI_VALUE_CONDITION_OPERATOR_UNSUPPORTED = (
        "multi_value_condition_operator_unsupported", "error"
    )
    MULTI_VALUE_MEMBERSHIP_ON_A_SINGLE_VALUE_COLUMN = (
        "multi_value_membership_on_a_single_value_column", "error"
    )
    MULTI_VALUE_SET_EQUALITY_UNSUPPORTED = (
        "multi_value_set_equality_unsupported", "error"
    )

    # --- schema-only rules, from validator.validate() ---
    AUTO_INCREMENT_PK_MUST_BE_ID = "auto_increment_pk_must_be_id", "error"
    COLUMN_NAME_TOO_LONG = "column_name_too_long", "error"
    CROSS_SITE_EXPANSION_UNHANDLED = "cross_site_expansion_unhandled", "error"
    DEFAULT_NOT_AN_ENUM_MEMBER = "default_not_an_enum_member", "error"
    DUPLICATE_COLUMN_NAME = "duplicate_column_name", "error"
    DUPLICATE_ENUM_MEMBER = "duplicate_enum_member", "error"
    DUPLICATE_ENUM_NAME = "duplicate_enum_name", "error"
    DUPLICATE_TABLE_NAME = "duplicate_table_name", "error"
    EMPTY_ENUM = "empty_enum", "warning"
    ILLEGAL_COLUMN_NAME_CHARACTER = "illegal_column_name_character", "error"
    LEGACY_CHOICE_TYPE = "legacy_choice_type", "error"
    # Arity, not type name. Both have a generic counterpart that would fire
    # instead -- UNIQUE_UNSUPPORTED_FOR_TYPE, and nothing at all for the
    # default -- and both are separate codes because the evidence and the
    # remedy are separate: the generic uniqueness message names the type
    # `'audit_event[]'`, which invites deleting the brackets rather than the
    # constraint, and there is no generic "default: is not supported" rule
    # for the second to widen. One declaration still produces one finding.
    MULTI_VALUE_DEFAULT_UNSUPPORTED = "multi_value_default_unsupported", "error"
    MULTI_VALUE_MEMBER_CONTAINS_THE_EXPORT_SEPARATOR = (
        "multi_value_member_contains_the_export_separator", "error",
    )
    MULTI_VALUE_UNIQUE_UNSUPPORTED = "multi_value_unique_unsupported", "error"
    ORPHAN_ENUM = "orphan_enum", "warning"
    RESERVED_COLUMN_NAME = "reserved_column_name", "error"
    UNIQUE_UNSUPPORTED_FOR_TYPE = "unique_unsupported_for_type", "error"
    UNIQUE_WITHOUT_NOT_NULL = "unique_without_not_null", "warning"
    UNKNOWN_COLUMN_TYPE = "unknown_column_type", "error"
    UNKNOWN_REF_TARGET = "unknown_ref_target", "error"



@dataclass(frozen=True, slots=True)
class Finding:
    """One thing the build has to say about the declaration it was given.

    `severity` is a property, not a field. It used to be the second
    constructor argument, which meant 155 call sites each restated how bad a
    rule is -- and the published catalogue restated it a 156th time. Nothing
    compared them, and one had already drifted.

    Deriving it from the code removes the disagreement rather than detecting
    it: there is no argument to get wrong, so no test is needed to catch a
    wrong one. A rule that genuinely comes in two strengths is two codes.
    """

    code: FindingCode
    message: str
    location: Location | None = None

    @property
    def severity(self) -> Severity:
        """How bad this is -- declared once, on the code."""
        return self.code.severity

    @property
    def detail(self) -> str:
        """The code and the message, for anything that shows a finding.

        The code is the identity and the published catalogue is keyed by it,
        but only the message used to reach the terminal and the manifest --
        and the message is prose that is free to be reworded in any commit.
        So the operator was shown the one part of a finding that is
        deliberately not searchable, and given nothing to carry them to the
        catalogue entry.

        One property rather than a format string at each site: the CLI's
        error path, the CLI's warning path and the manifest template all
        show findings, and three spellings of "how a finding looks" drift.
        The severity marker stays with the caller because it legitimately
        differs by medium -- `**[WARNING]**` in markdown, `[WARNING]` on a
        terminal.
        """
        return f"{self.code}: {self.message}"
