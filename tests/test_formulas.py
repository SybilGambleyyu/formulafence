from formulafence.formulas import (
    ParsedReference,
    StructuredTable,
    extract_references,
    formula_fingerprint,
    inspect_formula,
)


def test_fingerprint_normalises_relative_copy_patterns() -> None:
    assert formula_fingerprint("=A2*2", "B2") == formula_fingerprint("=A3*2", "B3")
    assert formula_fingerprint("=$A2*2", "B2") == formula_fingerprint("=$A3*2", "B3")
    assert formula_fingerprint("=A2*2", "B2") != formula_fingerprint("=A2*3", "B2")


def test_extract_references_keeps_external_workbooks_separate() -> None:
    references = extract_references("='Input Sheet'!$B$2+[book.xlsx]Sheet1!A1+SUM(C1:C3)")

    assert len(references) == 3
    assert references[0].sheet == "Input Sheet"
    assert references[0].min_column == 2
    assert references[0].min_row == 2
    assert references[1].is_external
    assert references[2].sheet is None
    assert references[2].is_range


def test_formula_inspection_resolves_names_and_marks_static_coverage_gaps() -> None:
    inspection = inspect_formula(
        '=HeadlineOutput+UnknownMetric+INDIRECT("Inputs!B2")',
        {
            "headlineoutput": (
                ParsedReference("Dashboard", 2, 12, 2, 12, raw="HeadlineOutput"),
            )
        },
    )

    assert inspection.references == (
        ParsedReference("Dashboard", 2, 12, 2, 12, raw="HeadlineOutput"),
    )
    assert inspection.unresolved_range_tokens == ("UnknownMetric",)
    assert inspection.dynamic_reference_functions == ("INDIRECT",)


def test_formula_inspection_resolves_static_structured_table_references() -> None:
    sales = StructuredTable(
        name="Sales",
        sheet="Data",
        ref="A1:C4",
        columns=("Amount", "Rate", "Value"),
        header_row_count=1,
        totals_row_count=0,
    )
    inspection = inspect_formula(
        "=SUM(Sales[Amount])+SUM(Sales[[#Data],[Amount]:[Rate]])+ROWS(Sales[#All])",
        structured_tables={"sales": sales},
    )

    assert inspection.references == (
        ParsedReference("Data", 1, 2, 1, 4, raw="Sales[Amount]"),
        ParsedReference("Data", 1, 2, 2, 4, raw="Sales[[#Data],[Amount]:[Rate]]"),
        ParsedReference("Data", 1, 1, 3, 4, raw="Sales[#All]"),
    )
    assert inspection.unresolved_range_tokens == ()


def test_formula_inspection_keeps_this_row_table_references_visible() -> None:
    sales = StructuredTable(
        name="Sales",
        sheet="Data",
        ref="A1:C4",
        columns=("Amount", "Rate", "Value"),
        header_row_count=1,
        totals_row_count=0,
    )

    inspection = inspect_formula("=Sales[@Amount]", structured_tables={"sales": sales})

    assert inspection.references == ()
    assert inspection.unresolved_range_tokens == ("Sales[@Amount]",)


def test_formula_inspection_resolves_table_header_data_and_total_regions() -> None:
    sales = StructuredTable(
        name="Sales",
        sheet="Data",
        ref="A1:C5",
        columns=("Amount", "Rate", "Value"),
        header_row_count=1,
        totals_row_count=1,
    )
    inspection = inspect_formula(
        "=SUM(Sales[[#Headers],[#Data],[Rate]])+SUM(Sales[[#Totals],[Value]])",
        structured_tables={"sales": sales},
    )

    assert inspection.references == (
        ParsedReference("Data", 2, 1, 2, 1, raw="Sales[[#Headers],[#Data],[Rate]]"),
        ParsedReference("Data", 2, 2, 2, 4, raw="Sales[[#Headers],[#Data],[Rate]]"),
        ParsedReference("Data", 3, 5, 3, 5, raw="Sales[[#Totals],[Value]]"),
    )
