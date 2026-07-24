from formulafence.formulas import extract_references, formula_fingerprint


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
