"""Safe, non-evaluating workbook inspection and dependency indexing."""

from __future__ import annotations

import hashlib
import warnings
from collections import defaultdict
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from openpyxl import load_workbook
from openpyxl.utils.cell import get_column_letter, range_boundaries
from openpyxl.utils.exceptions import InvalidFileException

from formulafence.formulas import (
    ParsedReference,
    StructuredTable,
    formula_fingerprint,
    has_broken_reference,
    inspect_formula,
    lambda_parameter_count,
    parse_reference_token,
    reference_lookup_key,
)
from formulafence.models import (
    CellKey,
    CellSnapshot,
    RangeDependency,
    SheetSnapshot,
    TableSnapshot,
    WorkbookLoadError,
    WorkbookSnapshot,
    display_location,
    json_safe_value,
)

_SUPPORTED_SUFFIXES = {".xlsx", ".xlsm"}
_CALCULATION_FIELDS = (
    "calcMode",
    "fullCalcOnLoad",
    "refMode",
    "iterate",
    "iterateCount",
    "iterateDelta",
    "fullPrecision",
    "calcCompleted",
    "calcOnSave",
    "concurrentCalc",
    "concurrentManualCount",
    "forceFullCalc",
)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file_handle:
        for block in iter(lambda: file_handle.read(1_048_576), b""):
            digest.update(block)
    return digest.hexdigest()


def _vba_hash(path: Path) -> str | None:
    """Hash the macro payload without loading or executing it."""
    try:
        with ZipFile(path) as archive:
            try:
                payload = archive.read("xl/vbaProject.bin")
            except KeyError:
                return None
    except BadZipFile as error:
        raise WorkbookLoadError(f"{path} is not a valid Office Open XML workbook") from error
    return hashlib.sha256(payload).hexdigest()


def _formula_text(value: object) -> str | None:
    if isinstance(value, str):
        return value
    # openpyxl represents array formulas as an object in recent versions.
    text = getattr(value, "text", None)
    return text if isinstance(text, str) else None


def _cell_snapshot(sheet: str, cell: object) -> CellSnapshot:
    coordinate = cell.coordinate
    value = cell.value
    data_type = cell.data_type
    formula = _formula_text(value) if data_type == "f" else None
    if formula is not None:
        return CellSnapshot(
            sheet=sheet,
            coordinate=coordinate,
            cell_type="formula",
            value=formula,
            value_type="formula",
            formula=formula,
            formula_fingerprint=formula_fingerprint(formula, coordinate),
        )
    cell_type = "error" if data_type == "e" else "value"
    return CellSnapshot(
        sheet=sheet,
        coordinate=coordinate,
        cell_type=cell_type,
        value=json_safe_value(value),
        value_type=type(value).__name__,
    )


def _calculation_settings(workbook: object) -> dict[str, object]:
    calculation = getattr(workbook, "calculation", None)
    if calculation is None:
        return {}
    return {
        field: value
        for field in _CALCULATION_FIELDS
        if (value := getattr(calculation, field, None)) is not None
    }


def _definition_text(definition: object) -> str:
    attr_text = getattr(definition, "attr_text", None)
    return str(attr_text if attr_text is not None else definition)


def _defined_names(workbook: object) -> dict[str, str]:
    """Inventory workbook and sheet-scoped names with unambiguous local keys."""
    names = getattr(workbook, "defined_names", {})
    result: dict[str, str] = {}
    try:
        items = names.items()
    except AttributeError:
        items = ()
    for name, definition in items:
        if getattr(definition, "localSheetId", None) is None:
            result[str(name)] = _definition_text(definition)
    for worksheet in getattr(workbook, "worksheets", ()):
        worksheet_names = getattr(worksheet, "defined_names", {})
        try:
            worksheet_items = worksheet_names.items()
        except AttributeError:
            continue
        for name, definition in worksheet_items:
            result[f"{worksheet.title}!{name}"] = _definition_text(definition)
    return result


def _named_destination_reference(
    name: str, sheet: str, coordinate: str
) -> ParsedReference | None:
    """Convert one workbook-defined-name destination into a static range."""
    parsed = parse_reference_token(f"{sheet}!{coordinate}")
    if parsed is None:
        return None
    return ParsedReference(
        parsed.sheet,
        parsed.min_column,
        parsed.min_row,
        parsed.max_column,
        parsed.max_row,
        raw=name,
        is_external=parsed.is_external,
    )


def _definition_references(name: str, definition: object) -> tuple[ParsedReference, ...]:
    """Return only static destinations from an ordinary defined name."""
    try:
        destinations = list(definition.destinations)
    except Exception:  # pragma: no cover - malformed name syntax is workbook-specific
        return ()
    return tuple(
        reference
        for sheet, coordinate in destinations
        if (reference := _named_destination_reference(name, sheet, coordinate)) is not None
    )


@dataclass(frozen=True)
class _FormulaDefinedName:
    """A defined name whose OOXML text is an Excel formula expression."""

    key: str
    formula: str
    scope: str | None


def _formula_defined_name(
    name: str, definition: object, scope: str | None
) -> _FormulaDefinedName | None:
    """Return a formula-valued defined name, excluding ordinary destinations.

    OOXML stores the text of a defined-name formula without its leading equals
    sign, while older writer output commonly retains one. ``DefinedName.type``
    lets us preserve ordinary A1 destinations and normalize either formula
    spelling before sending it through the formula tokenizer.
    """
    attr_text = getattr(definition, "attr_text", None)
    if not isinstance(attr_text, str):
        return None
    formula = attr_text.strip()
    if not formula:
        return None
    if not formula.startswith("="):
        try:
            definition_type = getattr(definition, "type", None)
        except Exception:  # pragma: no cover - malformed name text is workbook-specific
            definition_type = None
        if definition_type == "RANGE":
            return None
        formula = f"={formula}"
    return _FormulaDefinedName(reference_lookup_key(name), formula, scope)


def _qualified_name_key(sheet: str, name: str) -> str:
    """Build a lookup key for a sheet-local name without quote ambiguity."""
    escaped_sheet = sheet.replace("'", "''")
    return reference_lookup_key(f"'{escaped_sheet}'!{name}")


def _named_reference_maps(
    workbook: object,
    structured_tables: Mapping[str, StructuredTable],
    sheet_order: tuple[str, ...],
) -> tuple[
    dict[str, tuple[ParsedReference, ...]],
    dict[str, dict[str, tuple[ParsedReference, ...]]],
    dict[str, tuple[ParsedReference, ...] | None],
    dict[str, dict[str, tuple[ParsedReference, ...] | None]],
]:
    """Build direct, formula-defined, and callable named-LAMBDA maps.

    Formula-valued names are expanded only when every dependency is statically
    visible and internal. Relative references, dynamic functions, unresolved
    tokens, recursive LAMBDAs, external links, and 3-D spans remain unresolved
    at a use site instead of producing a guessed graph edge.
    """
    workbook_names = getattr(workbook, "defined_names", {})
    global_references: dict[str, tuple[ParsedReference, ...]] = {}
    global_formulas: dict[str, _FormulaDefinedName] = {}
    global_lambdas: dict[str, _FormulaDefinedName] = {}
    try:
        workbook_items = workbook_names.items()
    except AttributeError:
        workbook_items = ()
    for name, definition in workbook_items:
        if getattr(definition, "localSheetId", None) is not None:
            continue
        formula_definition = _formula_defined_name(str(name), definition, None)
        if formula_definition is not None:
            target = (
                global_lambdas
                if lambda_parameter_count(formula_definition.formula) is not None
                else global_formulas
            )
            target[formula_definition.key] = formula_definition
            continue
        references = _definition_references(str(name), definition)
        if references:
            global_references[reference_lookup_key(str(name))] = references

    local_references: dict[str, dict[str, tuple[ParsedReference, ...]]] = {}
    local_formulas: dict[str, dict[str, _FormulaDefinedName]] = {}
    local_lambdas: dict[str, dict[str, _FormulaDefinedName]] = {}
    sheet_titles: dict[str, str] = {}
    for worksheet in getattr(workbook, "worksheets", ()):
        scope = worksheet.title.casefold()
        sheet_titles[scope] = worksheet.title
        worksheet_names = getattr(worksheet, "defined_names", {})
        try:
            worksheet_items = worksheet_names.items()
        except AttributeError:
            continue
        sheet_references: dict[str, tuple[ParsedReference, ...]] = {}
        for name, definition in worksheet_items:
            formula_definition = _formula_defined_name(str(name), definition, scope)
            if formula_definition is not None:
                target = (
                    local_lambdas
                    if lambda_parameter_count(formula_definition.formula) is not None
                    else local_formulas
                )
                target.setdefault(scope, {})[formula_definition.key] = formula_definition
                continue
            references = _definition_references(str(name), definition)
            if not references:
                continue
            sheet_references[reference_lookup_key(str(name))] = references
        if sheet_references:
            local_references[scope] = sheet_references

    qualified_formulas = {
        _qualified_name_key(sheet_titles[scope], key): definition
        for scope, definitions in local_formulas.items()
        for key, definition in definitions.items()
    }
    qualified_lambdas = {
        _qualified_name_key(sheet_titles[scope], key): definition
        for scope, definitions in local_lambdas.items()
        for key, definition in definitions.items()
    }
    resolved_definitions: dict[tuple[str | None, str], tuple[ParsedReference, ...]] = {}
    resolving: set[tuple[str | None, str]] = set()
    failed: set[tuple[str | None, str]] = set()

    def identity_for(definition: _FormulaDefinedName) -> tuple[str | None, str]:
        return definition.scope, definition.key

    def cached_references(
        definition: _FormulaDefinedName,
    ) -> tuple[ParsedReference, ...] | None:
        return resolved_definitions.get(identity_for(definition))

    def has_cached_references(definition: _FormulaDefinedName) -> bool:
        return identity_for(definition) in resolved_definitions

    def visible_references(scope: str | None) -> dict[str, tuple[ParsedReference, ...]]:
        references = dict(global_references)
        for key, definition in global_formulas.items():
            if cached := cached_references(definition):
                references[key] = cached
            elif has_cached_references(definition):
                references[key] = ()
        for local_scope, values in local_references.items():
            for key, local_values in values.items():
                references[_qualified_name_key(sheet_titles[local_scope], key)] = local_values
        for local_scope, definitions in local_formulas.items():
            for key, definition in definitions.items():
                if cached := cached_references(definition):
                    references[_qualified_name_key(sheet_titles[local_scope], key)] = cached
                elif has_cached_references(definition):
                    references[_qualified_name_key(sheet_titles[local_scope], key)] = ()
        if scope is not None:
            references.update(local_references.get(scope, {}))
            for key, definition in local_formulas.get(scope, {}).items():
                if cached := cached_references(definition):
                    references[key] = cached
                elif has_cached_references(definition):
                    references[key] = ()
        return references

    def visible_named_functions(
        scope: str | None,
    ) -> dict[str, tuple[ParsedReference, ...] | None]:
        functions: dict[str, tuple[ParsedReference, ...] | None] = {
            key: cached_references(definition)
            for key, definition in global_lambdas.items()
        }
        for local_scope, definitions in local_lambdas.items():
            for key, definition in definitions.items():
                functions[_qualified_name_key(sheet_titles[local_scope], key)] = (
                    cached_references(definition)
                )
        if scope is not None:
            for key, definition in local_lambdas.get(scope, {}).items():
                functions[key] = cached_references(definition)
        return functions

    def named_definition_for(
        token: str, scope: str | None
    ) -> _FormulaDefinedName | None:
        key = reference_lookup_key(token)
        if qualified := qualified_formulas.get(key):
            return qualified
        if qualified := qualified_lambdas.get(key):
            return qualified
        if "!" in key:
            return None
        if scope is not None and (local := local_formulas.get(scope, {}).get(key)):
            return local
        if scope is not None and (local := local_lambdas.get(scope, {}).get(key)):
            return local
        return global_formulas.get(key) or global_lambdas.get(key)

    def resolve_definition(
        definition: _FormulaDefinedName,
    ) -> tuple[ParsedReference, ...] | None:
        identity = identity_for(definition)
        if has_cached_references(definition):
            return cached_references(definition)
        if identity in failed or identity in resolving:
            return None

        resolving.add(identity)
        resolved_references: tuple[ParsedReference, ...] | None = None
        try:
            inspection = inspect_formula(
                definition.formula,
                named_references=visible_references(definition.scope),
                structured_tables=structured_tables,
                sheet_order=sheet_order,
                named_function_references=visible_named_functions(definition.scope),
            )
            can_expand = not (
                has_broken_reference(definition.formula)
                or inspection.tokenization_failed
                or inspection.dynamic_reference_functions
                or inspection.three_d_reference_tokens
                or inspection.spill_reference_tokens
                or inspection.implicit_intersection_tokens
            )
            if can_expand:
                for token in inspection.unresolved_range_tokens:
                    dependency = named_definition_for(token, definition.scope)
                    if dependency is None or resolve_definition(dependency) is None:
                        can_expand = False
                        break
            if can_expand:
                inspection = inspect_formula(
                    definition.formula,
                    named_references=visible_references(definition.scope),
                    structured_tables=structured_tables,
                    sheet_order=sheet_order,
                    named_function_references=visible_named_functions(definition.scope),
                )
                if (
                    inspection.unresolved_range_tokens
                    or inspection.tokenization_failed
                    or inspection.dynamic_reference_functions
                    or inspection.three_d_reference_tokens
                    or inspection.spill_reference_tokens
                    or inspection.implicit_intersection_tokens
                    or any(
                        reference.is_external or reference.sheet is None
                        for reference in inspection.references
                    )
                ):
                    can_expand = False
            if can_expand:
                resolved_references = tuple(dict.fromkeys(inspection.references))
        finally:
            resolving.remove(identity)

        if resolved_references is None:
            failed.add(identity)
            return None
        resolved_definitions[identity] = resolved_references
        return resolved_references

    for definition in global_formulas.values():
        resolve_definition(definition)
    for definition in global_lambdas.values():
        resolve_definition(definition)
    for definitions in local_formulas.values():
        for definition in definitions.values():
            resolve_definition(definition)
    for definitions in local_lambdas.values():
        for definition in definitions.values():
            resolve_definition(definition)

    global_result = dict(global_references)
    for key, definition in global_formulas.items():
        if cached := cached_references(definition):
            global_result[key] = cached
        elif has_cached_references(definition):
            global_result[key] = ()
    for scope, values in local_references.items():
        for key, references in values.items():
            global_result[_qualified_name_key(sheet_titles[scope], key)] = references
    for scope, definitions in local_formulas.items():
        for key, definition in definitions.items():
            if cached := cached_references(definition):
                global_result[_qualified_name_key(sheet_titles[scope], key)] = cached
            elif has_cached_references(definition):
                global_result[_qualified_name_key(sheet_titles[scope], key)] = ()

    local_result: dict[str, dict[str, tuple[ParsedReference, ...]]] = {}
    for scope in set(local_references) | set(local_formulas):
        references = dict(local_references.get(scope, {}))
        for key, definition in local_formulas.get(scope, {}).items():
            if cached := cached_references(definition):
                references[key] = cached
            elif has_cached_references(definition):
                references[key] = ()
        if references:
            local_result[scope] = references

    global_function_result: dict[str, tuple[ParsedReference, ...] | None] = {
        key: cached_references(definition)
        for key, definition in global_lambdas.items()
    }
    for scope, definitions in local_lambdas.items():
        for key, definition in definitions.items():
            global_function_result[_qualified_name_key(sheet_titles[scope], key)] = (
                cached_references(definition)
            )

    local_function_result: dict[
        str, dict[str, tuple[ParsedReference, ...] | None]
    ] = {}
    for scope, definitions in local_lambdas.items():
        functions = {
            key: cached_references(definition) for key, definition in definitions.items()
        }
        if functions:
            local_function_result[scope] = functions
    return global_result, local_result, global_function_result, local_function_result


def _table_columns(
    worksheet: object,
    table: object,
    min_column: int,
    min_row: int,
    max_column: int,
    header_row_count: int,
) -> tuple[str, ...]:
    """Read table column labels, falling back to the inspectable header cells."""
    table_columns = tuple(getattr(table, "tableColumns", ()) or ())
    names = tuple(str(getattr(column, "name", "")) for column in table_columns)
    width = max_column - min_column + 1
    if len(names) == width and all(names):
        return names
    if header_row_count:
        return tuple(
            str(value) if (value := worksheet.cell(min_row, column).value) is not None else ""
            for column in range(min_column, max_column + 1)
        )
    return ()


def _table_snapshots(workbook: object) -> dict[str, TableSnapshot]:
    """Inventory Excel-table definitions that affect structured references."""
    result: dict[str, TableSnapshot] = {}
    for worksheet in getattr(workbook, "worksheets", ()):
        table_list = getattr(worksheet, "tables", {})
        try:
            table_values = table_list.values()
        except AttributeError:
            continue
        for table in table_values:
            name = str(
                getattr(table, "displayName", None)
                or getattr(table, "name", None)
                or ""
            )
            ref = getattr(table, "ref", None)
            if not name or not isinstance(ref, str):
                continue
            try:
                min_column, min_row, max_column, max_row = range_boundaries(ref)
            except ValueError:
                continue
            height = max_row - min_row + 1
            header_rows = min(max(int(getattr(table, "headerRowCount", 1) or 0), 0), height)
            totals_rows = min(
                max(int(getattr(table, "totalsRowCount", 0) or 0), 0), height - header_rows
            )
            result[name] = TableSnapshot(
                name=name,
                sheet=worksheet.title,
                ref=ref,
                columns=_table_columns(
                    worksheet,
                    table,
                    min_column,
                    min_row,
                    max_column,
                    header_rows,
                ),
                header_row_count=header_rows,
                totals_row_count=totals_rows,
            )
    return result


def _structured_table_map(tables: dict[str, TableSnapshot]) -> dict[str, StructuredTable]:
    """Translate stable table inventory records into formula-resolution metadata."""
    return {
        name.casefold(): StructuredTable(
            name=table.name,
            sheet=table.sheet,
            ref=table.ref,
            columns=table.columns,
            header_row_count=table.header_row_count,
            totals_row_count=table.totals_row_count,
        )
        for name, table in tables.items()
    }


def load_snapshot(path: str | Path) -> WorkbookSnapshot:
    """Load a workbook as a semantic snapshot without evaluating its contents."""
    source = Path(path)
    if not source.exists() or not source.is_file():
        raise WorkbookLoadError(f"Workbook does not exist or is not a file: {source}")
    if source.suffix.lower() not in _SUPPORTED_SUFFIXES:
        supported = ", ".join(sorted(_SUPPORTED_SUFFIXES))
        raise WorkbookLoadError(
            f"Unsupported workbook type {source.suffix!r}; supported types: {supported}"
        )

    try:
        with warnings.catch_warnings(record=True) as caught_warnings:
            warnings.simplefilter("always")
            workbook = load_workbook(
                source,
                read_only=False,
                data_only=False,
                keep_vba=False,
                keep_links=False,
                rich_text=False,
            )
    except (BadZipFile, InvalidFileException, OSError, ValueError) as error:
        raise WorkbookLoadError(f"Could not read workbook {source}: {error}") from error
    parser_warnings = tuple(sorted({str(warning.message) for warning in caught_warnings}))

    sheets: dict[str, SheetSnapshot] = {}
    cells: dict[CellKey, CellSnapshot] = {}
    reverse_dependencies: dict[CellKey, set[CellKey]] = defaultdict(set)
    range_dependencies: list[RangeDependency] = []
    external_references: set[CellKey] = set()
    broken_references: set[CellKey] = set()
    unresolved_reference_tokens: dict[CellKey, tuple[str, ...]] = {}
    dynamic_reference_functions: dict[CellKey, tuple[str, ...]] = {}
    three_d_reference_tokens: dict[CellKey, tuple[str, ...]] = {}
    spill_reference_tokens: dict[CellKey, tuple[str, ...]] = {}
    implicit_intersection_tokens: dict[CellKey, tuple[str, ...]] = {}
    tokenization_failure_cells: set[CellKey] = set()
    tables = _table_snapshots(workbook)
    structured_tables = _structured_table_map(tables)
    sheet_order = tuple(worksheet.title for worksheet in workbook.worksheets)
    (
        global_named_references,
        local_named_references,
        global_named_functions,
        local_named_functions,
    ) = _named_reference_maps(workbook, structured_tables, sheet_order)

    for worksheet in workbook.worksheets:
        named_references = {
            **global_named_references,
            **local_named_references.get(worksheet.title.casefold(), {}),
        }
        named_functions = {
            **global_named_functions,
            **local_named_functions.get(worksheet.title.casefold(), {}),
        }
        nonempty_cells = 0
        formula_cells = 0
        # _cells lets us avoid traversing a sheet's whole used rectangle when a
        # workbook has a sparse, accidentally enormous dimension.
        worksheet_cells = sorted(
            worksheet._cells.values(),  # noqa: SLF001 - needed for sparse workbook safety
            key=lambda current: (current.row, current.column),
        )
        for cell in worksheet_cells:
            if cell.value is None:
                continue
            snapshot = _cell_snapshot(worksheet.title, cell)
            cells[snapshot.location] = snapshot
            nonempty_cells += 1
            if not snapshot.is_formula or snapshot.formula is None:
                continue

            formula_cells += 1
            if has_broken_reference(snapshot.formula):
                broken_references.add(snapshot.location)
            inspection = inspect_formula(
                snapshot.formula,
                named_references=named_references,
                structured_tables=structured_tables,
                origin=snapshot.location,
                sheet_order=sheet_order,
                named_function_references=named_functions,
            )
            if inspection.unresolved_range_tokens:
                unresolved_reference_tokens[snapshot.location] = inspection.unresolved_range_tokens
            if inspection.tokenization_failed:
                tokenization_failure_cells.add(snapshot.location)
            if inspection.dynamic_reference_functions:
                dynamic_reference_functions[snapshot.location] = (
                    inspection.dynamic_reference_functions
                )
            if inspection.three_d_reference_tokens:
                three_d_reference_tokens[snapshot.location] = inspection.three_d_reference_tokens
            if inspection.spill_reference_tokens:
                spill_reference_tokens[snapshot.location] = inspection.spill_reference_tokens
            if inspection.implicit_intersection_tokens:
                implicit_intersection_tokens[snapshot.location] = (
                    inspection.implicit_intersection_tokens
                )
            for reference in inspection.references:
                if reference.is_external:
                    external_references.add(snapshot.location)
                    continue
                if None in {
                    reference.min_column,
                    reference.min_row,
                    reference.max_column,
                    reference.max_row,
                }:
                    continue
                source_sheet = reference.sheet or worksheet.title
                if reference.is_range:
                    range_dependencies.append(
                        RangeDependency(
                            source_sheet=source_sheet,
                            min_column=reference.min_column,
                            min_row=reference.min_row,
                            max_column=reference.max_column,
                            max_row=reference.max_row,
                            dependent=snapshot.location,
                        )
                    )
                else:
                    source_coordinate = (
                        f"{get_column_letter(reference.min_column)}{reference.min_row}"
                    )
                    reverse_dependencies[(source_sheet, source_coordinate)].add(snapshot.location)

        sheets[worksheet.title] = SheetSnapshot(
            title=worksheet.title,
            state=worksheet.sheet_state,
            nonempty_cells=nonempty_cells,
            formula_cells=formula_cells,
            max_row=worksheet.max_row,
            max_column=worksheet.max_column,
        )

    return WorkbookSnapshot(
        path=source,
        sha256=sha256_file(source),
        file_type=source.suffix.lower().lstrip("."),
        sheets=sheets,
        cells=cells,
        reverse_dependencies=dict(reverse_dependencies),
        range_dependencies=range_dependencies,
        external_references=external_references,
        broken_references=broken_references,
        unresolved_reference_tokens=unresolved_reference_tokens,
        dynamic_reference_functions=dynamic_reference_functions,
        three_d_reference_tokens=three_d_reference_tokens,
        spill_reference_tokens=spill_reference_tokens,
        implicit_intersection_tokens=implicit_intersection_tokens,
        tokenization_failure_cells=tokenization_failure_cells,
        tables=tables,
        sheet_order=sheet_order,
        defined_names=_defined_names(workbook),
        macro_hash=_vba_hash(source),
        calculation_settings=_calculation_settings(workbook),
        parser_warnings=parser_warnings,
    )


def profile_snapshot(snapshot: WorkbookSnapshot) -> dict[str, object]:
    """Return a data-minimising inventory suitable for a safe review artifact."""
    return {
        "schema_version": "1.0",
        "workbook": snapshot.summary(),
        "sheets": [sheet.to_dict() for sheet in snapshot.sheets.values()],
        "tables": [
            snapshot.tables[name].to_dict()
            for name in sorted(snapshot.tables, key=str.casefold)
        ],
        "defined_names": sorted(snapshot.defined_names),
        "calculation_settings": snapshot.calculation_settings,
        "features": {
            "external_reference_cells": [
                f"{sheet}!{coordinate}"
                for sheet, coordinate in sorted(snapshot.external_references)
            ],
            "broken_reference_cells": [
                f"{sheet}!{coordinate}"
                for sheet, coordinate in sorted(snapshot.broken_references)
            ],
            "has_vba": snapshot.macro_hash is not None,
            "parser_warnings": list(snapshot.parser_warnings),
            "unresolved_reference_cells": [
                {
                    "location": display_location(location),
                    "tokens": list(tokens),
                }
                for location, tokens in sorted(snapshot.unresolved_reference_tokens.items())
            ],
            "dynamic_reference_cells": [
                {
                    "location": display_location(location),
                    "functions": list(functions),
                }
                for location, functions in sorted(snapshot.dynamic_reference_functions.items())
            ],
            "three_d_reference_cells": [
                {
                    "location": display_location(location),
                    "tokens": list(tokens),
                }
                for location, tokens in sorted(snapshot.three_d_reference_tokens.items())
            ],
            "spill_reference_cells": [
                {
                    "location": display_location(location),
                    "tokens": list(tokens),
                }
                for location, tokens in sorted(snapshot.spill_reference_tokens.items())
            ],
            "implicit_intersection_cells": [
                {
                    "location": display_location(location),
                    "tokens": list(tokens),
                }
                for location, tokens in sorted(
                    snapshot.implicit_intersection_tokens.items()
                )
            ],
            "tokenization_failure_cells": [
                display_location(location)
                for location in sorted(snapshot.tokenization_failure_cells)
            ],
        },
    }
