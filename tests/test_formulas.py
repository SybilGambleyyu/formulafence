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


def test_formula_inspection_recognises_a_known_named_constant() -> None:
    inspection = inspect_formula("=StaticRate", {"staticrate": ()})

    assert inspection.references == ()
    assert inspection.unresolved_range_tokens == ()


def test_formula_inspection_marks_tokenization_failures() -> None:
    inspection = inspect_formula("=SUM(A1#)")

    assert inspection.tokenization_failed is True


def test_formula_inspection_respects_let_lexical_scope() -> None:
    inspection = inspect_formula(
        "=LET(rate,ExistingRate,LET(rate,Inputs!B3,rate)+rate+UnknownMetric)",
        {
            "existingrate": (
                ParsedReference("Inputs", 2, 2, 2, 2, raw="ExistingRate"),
            )
        },
    )

    assert inspection.references == (
        ParsedReference("Inputs", 2, 2, 2, 2, raw="ExistingRate"),
        ParsedReference("Inputs", 2, 3, 2, 3, raw="Inputs!B3"),
    )
    assert inspection.unresolved_range_tokens == ("UnknownMetric",)


def test_formula_inspection_handles_microsoft_let_example_without_false_gaps() -> None:
    inspection = inspect_formula(
        '=LET(filterCriteria,"Fred",filteredRange,FILTER(A2:D8,A2:A8=filterCriteria),'
        'IF(ISBLANK(filteredRange),"-",filteredRange))'
    )

    assert inspection.references == (
        ParsedReference(None, 1, 2, 4, 8, raw="A2:D8"),
        ParsedReference(None, 1, 2, 1, 8, raw="A2:A8"),
    )
    assert inspection.unresolved_range_tokens == ()


def test_formula_inspection_does_not_shadow_a_let_value_before_its_binding() -> None:
    inspection = inspect_formula(
        "=LET(rate,rate+Inputs!B2,rate*2)",
        {"rate": (ParsedReference("Inputs", 2, 3, 2, 3, raw="rate"),)},
    )

    assert inspection.references == (
        ParsedReference("Inputs", 2, 3, 2, 3, raw="rate"),
        ParsedReference("Inputs", 2, 2, 2, 2, raw="Inputs!B2"),
    )
    assert inspection.unresolved_range_tokens == ()


def test_formula_inspection_respects_inline_lambda_parameters() -> None:
    lambda_call = inspect_formula("=LAMBDA(rate,rate*Inputs!B2)(Inputs!B3)")
    reduce_call = inspect_formula(
        "=REDUCE(0,Inputs!B2:B3,LAMBDA(acc,value,acc+value))"
    )
    namespaced_let = inspect_formula("=_xlfn.LET(rate,Inputs!B2,rate*2)")

    assert lambda_call.references == (
        ParsedReference("Inputs", 2, 2, 2, 2, raw="Inputs!B2"),
        ParsedReference("Inputs", 2, 3, 2, 3, raw="Inputs!B3"),
    )
    assert lambda_call.unresolved_range_tokens == ()
    assert reduce_call.references == (
        ParsedReference("Inputs", 2, 2, 2, 3, raw="Inputs!B2:B3"),
    )
    assert reduce_call.unresolved_range_tokens == ()
    assert namespaced_let.unresolved_range_tokens == ()


def test_formula_inspection_never_treats_a_cell_reference_as_a_local_variable() -> None:
    inspection = inspect_formula("=LET(A1,Inputs!B2,A1)")

    assert [reference.raw for reference in inspection.references] == [
        "A1",
        "Inputs!B2",
        "A1",
    ]


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


def test_formula_inspection_resolves_three_d_references_in_tab_order() -> None:
    inspection = inspect_formula(
        "=SUM('Jan 2026:Mar 2026'!$B$2:$B$3)",
        sheet_order=("Jan 2026", "Feb 2026", "Mar 2026", "Summary"),
    )

    assert inspection.references == (
        ParsedReference("Jan 2026", 2, 2, 2, 3, raw="'Jan 2026:Mar 2026'!$B$2:$B$3"),
        ParsedReference("Feb 2026", 2, 2, 2, 3, raw="'Jan 2026:Mar 2026'!$B$2:$B$3"),
        ParsedReference("Mar 2026", 2, 2, 2, 3, raw="'Jan 2026:Mar 2026'!$B$2:$B$3"),
    )
    assert inspection.three_d_reference_tokens == ("'Jan 2026:Mar 2026'!$B$2:$B$3",)
    assert inspection.unresolved_range_tokens == ()


def test_formula_inspection_keeps_three_d_references_visible_without_sheet_order() -> None:
    inspection = inspect_formula("=SUM(Jan:Mar!B2)")
    missing_endpoint = inspect_formula(
        "=SUM(Jan:Mar!B2)", sheet_order=("Jan", "Summary")
    )
    external = inspect_formula(
        "=SUM('[book.xlsx]Jan:Mar'!B2)", sheet_order=("Jan", "Feb", "Mar")
    )

    assert inspection.references == ()
    assert inspection.three_d_reference_tokens == ()
    assert inspection.unresolved_range_tokens == ("Jan:Mar!B2",)
    assert missing_endpoint.unresolved_range_tokens == ("Jan:Mar!B2",)
    assert external.references[0].is_external
    assert external.unresolved_range_tokens == ()


def test_formula_inspection_requires_table_data_origin_for_this_row_references() -> None:
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


def test_formula_inspection_resolves_current_row_table_references_in_context() -> None:
    sales = StructuredTable(
        name="Sales",
        sheet="Data",
        ref="A1:C4",
        columns=("Sales Amount", "Rate", "Value"),
        header_row_count=1,
        totals_row_count=0,
    )

    inspection = inspect_formula(
        "=[@[Sales Amount]]+[@Rate]+[Sales Amount]+[[Sales Amount]:[Rate]]"
        "+Sales[@Rate]+Sales[[#This Row],[Sales Amount]]"
        "+Sales[[#This Row],[Sales Amount]:[Rate]]",
        structured_tables={"sales": sales},
        origin=("Data", "C3"),
    )

    assert inspection.references == (
        ParsedReference("Data", 1, 3, 1, 3, raw="[@[Sales Amount]]"),
        ParsedReference("Data", 2, 3, 2, 3, raw="[@Rate]"),
        ParsedReference("Data", 1, 3, 1, 3, raw="[Sales Amount]"),
        ParsedReference("Data", 1, 3, 2, 3, raw="[[Sales Amount]:[Rate]]"),
        ParsedReference("Data", 2, 3, 2, 3, raw="Sales[@Rate]"),
        ParsedReference(
            "Data", 1, 3, 1, 3, raw="Sales[[#This Row],[Sales Amount]]"
        ),
        ParsedReference(
            "Data", 1, 3, 2, 3, raw="Sales[[#This Row],[Sales Amount]:[Rate]]"
        ),
    )
    assert inspection.unresolved_range_tokens == ()


def test_formula_inspection_keeps_invalid_current_row_references_visible() -> None:
    sales = StructuredTable(
        name="Sales",
        sheet="Data",
        ref="A1:C5",
        columns=("Amount", "Rate", "Value"),
        header_row_count=1,
        totals_row_count=1,
    )

    outside = inspect_formula(
        "=Sales[@Amount]", structured_tables={"sales": sales}, origin=("Report", "B2")
    )
    header = inspect_formula(
        "=[@Amount]", structured_tables={"sales": sales}, origin=("Data", "A1")
    )
    adjacent_unqualified = inspect_formula(
        "=[Amount]", structured_tables={"sales": sales}, origin=("Data", "D3")
    )
    total = inspect_formula(
        "=Sales[[#This Row],[Amount]]",
        structured_tables={"sales": sales},
        origin=("Data", "C5"),
    )

    assert outside.unresolved_range_tokens == ("Sales[@Amount]",)
    assert header.unresolved_range_tokens == ("[@Amount]",)
    assert adjacent_unqualified.unresolved_range_tokens == ("[Amount]",)
    assert total.unresolved_range_tokens == ("Sales[[#This Row],[Amount]]",)


def test_formula_inspection_resolves_qualified_this_row_from_an_adjacent_cell() -> None:
    sales = StructuredTable(
        name="Sales",
        sheet="Data",
        ref="A1:C5",
        columns=("Amount", "Rate", "Value"),
        header_row_count=1,
        totals_row_count=1,
    )

    inspection = inspect_formula(
        "=Sales[@Amount]+Sales[[#This Row],[Amount]:[Rate]]",
        structured_tables={"sales": sales},
        origin=("Data", "D3"),
    )

    assert inspection.references == (
        ParsedReference("Data", 1, 3, 1, 3, raw="Sales[@Amount]"),
        ParsedReference(
            "Data", 1, 3, 2, 3, raw="Sales[[#This Row],[Amount]:[Rate]]"
        ),
    )
    assert inspection.unresolved_range_tokens == ()


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
