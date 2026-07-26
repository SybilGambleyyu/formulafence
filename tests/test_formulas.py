from formulafence.formulas import (
    ParsedReference,
    StructuredTable,
    extract_references,
    formula_fingerprint,
    inspect_formula,
    lambda_parameter_count,
)


def test_fingerprint_normalises_relative_copy_patterns() -> None:
    assert formula_fingerprint("=A2*2", "B2") == formula_fingerprint("=A3*2", "B3")
    assert formula_fingerprint("=$A2*2", "B2") == formula_fingerprint("=$A3*2", "B3")
    assert formula_fingerprint("=A2*2", "B2") != formula_fingerprint("=A2*3", "B2")
    assert formula_fingerprint("=A1#", "B1") == formula_fingerprint("=A2#", "B2")
    assert formula_fingerprint("=A1#", "B1") == formula_fingerprint(
        "=_xlfn.ANCHORARRAY(A1)", "B1"
    )
    assert formula_fingerprint("=A1#", "B1") != formula_fingerprint("=A1", "B1")
    assert formula_fingerprint("=@A1:A3", "B2") == formula_fingerprint(
        "=@A2:A4", "B3"
    )
    assert formula_fingerprint("=@A1:A3", "B2") != formula_fingerprint("=A1:A3", "B2")
    assert formula_fingerprint("=@A1:A3", "B2") == formula_fingerprint(
        "=_xlfn.SINGLE(A1:A3)", "B2"
    )


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


def test_formula_inspection_inventories_external_action_and_provider_functions() -> None:
    inspection = inspect_formula(
        '=_xlfn.IMAGE("https://private.example.test/image.png")'
        '&HYPERLINK("#Inputs!A1","Internal")'
        '&WEBSERVICE("https://private.example.test/service")'
        '&RTD("Private.Provider","PrivateServer","Topic")'
        '&HYPERLINK("https://private.example.test/second","External")'
        '+_xlfn.STOCKHISTORY("XNAS:MSFT",DATE(2024,1,1))'
        '+CUBEVALUE("Finance","[Measures].[Revenue]")'
        '+CUBEMEMBER("Finance","[Date].[All]")'
        '+CUBEMEMBERPROPERTY("Finance","[Date].[All]","MEMBER_CAPTION")'
        '+CUBERANKEDMEMBER("Finance",A1,1)'
        '+CUBESET("Finance","[Date].[All].Children")'
        '+CUBESETCOUNT(A1)'
        '+CUBEKPIMEMBER("Finance","KPI",1)'
    )
    shadowed = inspect_formula(
        '=STOCKHISTORY(A1,DATE(2024,1,1))',
        named_function_references={"stockhistory": ()},
    )

    assert inspection.external_action_functions == (
        "IMAGE",
        "HYPERLINK",
        "WEBSERVICE",
        "RTD",
        "HYPERLINK",
        "STOCKHISTORY",
        "CUBEVALUE",
        "CUBEMEMBER",
        "CUBEMEMBERPROPERTY",
        "CUBERANKEDMEMBER",
        "CUBESET",
        "CUBESETCOUNT",
        "CUBEKPIMEMBER",
    )
    assert shadowed.external_action_functions == ()


def test_formula_inspection_propagates_external_actions_from_named_definitions() -> None:
    named_lambda = inspect_formula(
        "=FENCE.WRAPPER(A1)",
        named_function_references={"fence.wrapper": ()},
        named_function_formula_external_action_functions={
            "fence.wrapper": ("HYPERLINK",),
        },
    )
    named_formula = inspect_formula(
        "=FENCE.DIRECT",
        named_references={"fence.direct": ()},
        named_formula_external_action_functions={
            "fence.direct": ("WEBSERVICE",),
        },
    )
    named_provider = inspect_formula(
        "=FENCE.MARKET(A1)",
        named_function_references={"fence.market": ()},
        named_function_formula_external_action_functions={
            "fence.market": ("STOCKHISTORY", "CUBEVALUE"),
        },
    )

    assert named_lambda.external_action_functions == ("HYPERLINK",)
    assert named_lambda.references == (
        ParsedReference(None, 1, 1, 1, 1, raw="A1"),
    )
    assert named_formula.external_action_functions == ("WEBSERVICE",)
    assert named_formula.unresolved_range_tokens == ()
    assert named_provider.external_action_functions == ("STOCKHISTORY", "CUBEVALUE")


def test_formula_inspection_inventories_python_in_excel_function_spellings() -> None:
    inspection = inspect_formula(
        "=_xlfn._xlws.PY(0,0,A1)+_xlws.PY(1,1,A2)+PY(2,0,A3)",
        origin=("Inputs", "B2"),
    )

    assert inspection.python_functions == ("PY", "PY", "PY")
    assert inspection.references == (
        ParsedReference(None, 1, 1, 1, 1, raw="A1"),
        ParsedReference(None, 1, 2, 1, 2, raw="A2"),
        ParsedReference(None, 1, 3, 1, 3, raw="A3"),
    )


def test_formula_inspection_inventories_namespaced_custom_function_candidates() -> None:
    inspection = inspect_formula(
        "=CONTOSO.ADD(A1)+CONTOSO.GETPRICE(A2)+MYFUNCTION.SPHEREVOLUME(A3)"
        "+ECMA.CEILING(A4,1)+ERF.PRECISE(A5)+WORKDAY.INTL(A6,1)"
        "+_xlfn._xlws.PY(0,0,A7)"
    )
    named_lambda = inspect_formula(
        "=LOCAL.RATE(A1)",
        named_function_references={"local.rate": ()},
    )
    named_formula = inspect_formula(
        "=MODEL.RATE(A1)",
        named_references={"model.rate": ()},
    )

    assert inspection.office_custom_function_candidates == (
        "CONTOSO.ADD",
        "CONTOSO.GETPRICE",
        "MYFUNCTION.SPHEREVOLUME",
    )
    assert named_lambda.office_custom_function_candidates == ()
    assert named_formula.office_custom_function_candidates == ()


def test_formula_inspection_propagates_candidates_from_named_definitions() -> None:
    named_lambda = inspect_formula(
        "=FENCE.WRAPPER(A1)",
        named_function_references={"fence.wrapper": ()},
        named_function_custom_function_candidates={
            "fence.wrapper": ("CONTOSO.GETDATA",),
        },
    )
    named_formula = inspect_formula(
        "=FENCE.DIRECT",
        named_references={"fence.direct": ()},
        named_custom_function_candidates={
            "fence.direct": ("CONTOSO.GETDATA",),
        },
    )

    assert named_lambda.office_custom_function_candidates == ("CONTOSO.GETDATA",)
    assert named_lambda.references == (
        ParsedReference(None, 1, 1, 1, 1, raw="A1"),
    )
    assert named_formula.office_custom_function_candidates == ("CONTOSO.GETDATA",)
    assert named_formula.unresolved_range_tokens == ()


def test_formula_inspection_inventories_worksheet_code_resource_registrations() -> None:
    inspection = inspect_formula(
        '=REGISTER.ID(A1,A2,"J!")+@REGISTER.ID(A3,A4,"J!")+ECMA.CEILING(A5,1)'
    )
    shadowed = inspect_formula(
        "=REGISTER.ID(A1,A2)",
        named_function_references={"register.id": ()},
    )

    assert inspection.worksheet_code_resource_registration_functions == (
        "REGISTER.ID",
        "REGISTER.ID",
    )
    assert inspection.office_custom_function_candidates == ()
    assert shadowed.worksheet_code_resource_registration_functions == ()


def test_formula_inspection_propagates_worksheet_code_resource_registrations() -> None:
    named_lambda = inspect_formula(
        "=FENCE.REGISTER(A1,A2)",
        named_function_references={"fence.register": ()},
        named_function_worksheet_code_resource_registration_functions={
            "fence.register": ("REGISTER.ID",)
        },
    )
    named_formula = inspect_formula(
        "=FENCE.DIRECT",
        named_references={"fence.direct": ()},
        named_worksheet_code_resource_registration_functions={
            "fence.direct": ("REGISTER.ID",)
        },
    )

    assert named_lambda.worksheet_code_resource_registration_functions == (
        "REGISTER.ID",
    )
    assert named_lambda.references == (
        ParsedReference(None, 1, 1, 1, 1, raw="A1"),
        ParsedReference(None, 1, 2, 1, 2, raw="A2"),
    )
    assert named_formula.worksheet_code_resource_registration_functions == (
        "REGISTER.ID",
    )
    assert named_formula.unresolved_range_tokens == ()


def test_formula_inspection_inventories_formula_defined_xlm_registrations() -> None:
    ordinary = inspect_formula('=REGISTER(A1,A2,"J!")+@REGISTER(A3,A4,"J!")')
    definition = inspect_formula(
        '=REGISTER(A1,A2,"J!")+@REGISTER(A3,A4,"J!")',
        inspect_formula_defined_xlm_registrations=True,
    )
    shadowed = inspect_formula(
        '=REGISTER(A1,A2,"J!")',
        named_function_references={"register": ()},
        inspect_formula_defined_xlm_registrations=True,
    )

    assert ordinary.formula_defined_xlm_registration_functions == ()
    assert definition.formula_defined_xlm_registration_functions == (
        "REGISTER",
        "REGISTER",
    )
    assert shadowed.formula_defined_xlm_registration_functions == ()


def test_formula_inspection_propagates_formula_defined_xlm_registrations() -> None:
    named_lambda = inspect_formula(
        "=FENCE.REGISTER(A1,A2)",
        named_function_references={"fence.register": ()},
        named_function_formula_defined_xlm_registration_functions={
            "fence.register": ("REGISTER",)
        },
    )
    named_formula = inspect_formula(
        "=FENCE.DIRECT",
        named_references={"fence.direct": ()},
        named_formula_defined_xlm_registration_functions={
            "fence.direct": ("REGISTER",)
        },
    )

    assert named_lambda.formula_defined_xlm_registration_functions == ("REGISTER",)
    assert named_lambda.references == (
        ParsedReference(None, 1, 1, 1, 1, raw="A1"),
        ParsedReference(None, 1, 2, 1, 2, raw="A2"),
    )
    assert named_formula.formula_defined_xlm_registration_functions == ("REGISTER",)
    assert named_formula.unresolved_range_tokens == ()


def test_formula_inspection_inventories_formula_defined_xlm_evaluations() -> None:
    ordinary = inspect_formula('=EVALUATE(A1)+@EVALUATE(A2)')
    definition = inspect_formula(
        '=EVALUATE(A1)+@EVALUATE(A2)',
        inspect_formula_defined_xlm_evaluations=True,
    )
    shadowed = inspect_formula(
        "=EVALUATE(A1)",
        named_function_references={"evaluate": ()},
        inspect_formula_defined_xlm_evaluations=True,
    )

    assert ordinary.formula_defined_xlm_evaluation_functions == ()
    assert definition.formula_defined_xlm_evaluation_functions == (
        "EVALUATE",
        "EVALUATE",
    )
    assert shadowed.formula_defined_xlm_evaluation_functions == ()


def test_formula_inspection_propagates_formula_defined_xlm_evaluations() -> None:
    named_lambda = inspect_formula(
        "=FENCE.EVALUATE(A1)",
        named_function_references={"fence.evaluate": ()},
        named_function_formula_defined_xlm_evaluation_functions={
            "fence.evaluate": ("EVALUATE",)
        },
    )
    named_formula = inspect_formula(
        "=FENCE.DIRECT",
        named_references={"fence.direct": ()},
        named_formula_defined_xlm_evaluation_functions={
            "fence.direct": ("EVALUATE",)
        },
    )

    assert named_lambda.formula_defined_xlm_evaluation_functions == ("EVALUATE",)
    assert named_lambda.references == (
        ParsedReference(None, 1, 1, 1, 1, raw="A1"),
    )
    assert named_formula.formula_defined_xlm_evaluation_functions == ("EVALUATE",)
    assert named_formula.unresolved_range_tokens == ()


def test_formula_inspection_inventories_formula_defined_xlm_get_cell_calls() -> None:
    ordinary = inspect_formula("=GET.CELL(7,A1)+@GET.CELL(53,A2)")
    definition = inspect_formula(
        "=GET.CELL(7,A1)+@GET.CELL(53,A2)",
        inspect_formula_defined_xlm_get_cell_calls=True,
    )
    shadowed = inspect_formula(
        "=GET.CELL(7,A1)",
        named_function_references={"get.cell": ()},
        inspect_formula_defined_xlm_get_cell_calls=True,
    )

    assert ordinary.formula_defined_xlm_get_cell_functions == ()
    assert ordinary.office_custom_function_candidates == ()
    assert definition.formula_defined_xlm_get_cell_functions == (
        "GET.CELL",
        "GET.CELL",
    )
    assert shadowed.formula_defined_xlm_get_cell_functions == ()


def test_formula_inspection_propagates_formula_defined_xlm_get_cell_calls() -> None:
    named_lambda = inspect_formula(
        "=FENCE.GET.CELL(A1)",
        named_function_references={"fence.get.cell": ()},
        named_function_formula_defined_xlm_get_cell_functions={
            "fence.get.cell": ("GET.CELL",)
        },
    )
    named_formula = inspect_formula(
        "=FENCE.DIRECT",
        named_references={"fence.direct": ()},
        named_formula_defined_xlm_get_cell_functions={
            "fence.direct": ("GET.CELL",)
        },
    )

    assert named_lambda.formula_defined_xlm_get_cell_functions == ("GET.CELL",)
    assert named_lambda.references == (
        ParsedReference(None, 1, 1, 1, 1, raw="A1"),
    )
    assert named_formula.formula_defined_xlm_get_cell_functions == ("GET.CELL",)
    assert named_formula.unresolved_range_tokens == ()


def test_formula_inspection_inventories_xlm_environment_information_calls() -> None:
    ordinary = inspect_formula(
        "=GET.WORKBOOK(4)+@GET.WORKSPACE(2)+GET.DOCUMENT(37)"
    )
    definition = inspect_formula(
        "=GET.WORKBOOK(4)+@GET.WORKSPACE(2)+GET.DOCUMENT(37)",
        inspect_formula_defined_xlm_environment_information_calls=True,
    )
    shadowed = inspect_formula(
        "=GET.WORKBOOK(4)",
        named_function_references={"get.workbook": ()},
        inspect_formula_defined_xlm_environment_information_calls=True,
    )

    assert ordinary.formula_defined_xlm_environment_information_functions == ()
    assert ordinary.office_custom_function_candidates == ()
    assert definition.formula_defined_xlm_environment_information_functions == (
        "GET.WORKBOOK",
        "GET.WORKSPACE",
        "GET.DOCUMENT",
    )
    assert shadowed.formula_defined_xlm_environment_information_functions == ()


def test_formula_inspection_propagates_xlm_environment_information_calls() -> None:
    named_lambda = inspect_formula(
        "=FENCE.ENVIRONMENT(A1)",
        named_function_references={"fence.environment": ()},
        named_function_formula_defined_xlm_environment_information_functions={
            "fence.environment": ("GET.WORKBOOK", "GET.WORKSPACE")
        },
    )
    named_formula = inspect_formula(
        "=FENCE.DIRECT",
        named_references={"fence.direct": ()},
        named_formula_defined_xlm_environment_information_functions={
            "fence.direct": ("GET.DOCUMENT",)
        },
    )

    assert named_lambda.formula_defined_xlm_environment_information_functions == (
        "GET.WORKBOOK",
        "GET.WORKSPACE",
    )
    assert named_lambda.references == (
        ParsedReference(None, 1, 1, 1, 1, raw="A1"),
    )
    assert named_formula.formula_defined_xlm_environment_information_functions == (
        "GET.DOCUMENT",
    )
    assert named_formula.unresolved_range_tokens == ()


def test_formula_inspection_inventories_native_environment_information_calls() -> None:
    ordinary = inspect_formula(
        '=CELL("filename")+@CELL("type",A1)+INFO("directory")'
    )
    shadowed = inspect_formula(
        '=CELL("filename")',
        named_function_references={"cell": ()},
    )

    assert ordinary.formula_environment_information_functions == (
        "CELL",
        "CELL",
        "INFO",
    )
    assert ordinary.formula_environment_information_implicit_cell_reference_count == 1
    assert ordinary.formula_environment_information_function_count == 3
    assert shadowed.formula_environment_information_functions == ()


def test_formula_inspection_propagates_native_environment_information_calls() -> None:
    named_lambda = inspect_formula(
        "=FENCE.ENVIRONMENT(A1)",
        named_function_references={"fence.environment": ()},
        named_function_formula_environment_information_functions={
            "fence.environment": ("CELL", "INFO")
        },
    )
    named_formula = inspect_formula(
        "=FENCE.DIRECT",
        named_references={"fence.direct": ()},
        named_formula_environment_information_functions={
            "fence.direct": ("CELL",)
        },
    )

    assert named_lambda.formula_environment_information_functions == ("CELL", "INFO")
    assert named_lambda.references == (
        ParsedReference(None, 1, 1, 1, 1, raw="A1"),
    )
    assert named_formula.formula_environment_information_functions == ("CELL",)
    assert named_formula.unresolved_range_tokens == ()


def test_formula_inspection_recognises_a_known_named_constant() -> None:
    inspection = inspect_formula("=StaticRate", {"staticrate": ()})

    assert inspection.references == ()
    assert inspection.unresolved_range_tokens == ()


def test_formula_inspection_traces_static_spill_anchors_and_marks_the_boundary() -> None:
    literal = inspect_formula("=SUM(A1#)")
    qualified = inspect_formula("=SUM('My Sheet'!$B$2#)")
    serialized = inspect_formula("=SUM(_xlfn.ANCHORARRAY(A1))")
    string_literal = inspect_formula('=CONCAT("literal A1#",A1#)')

    assert literal.references == (ParsedReference(None, 1, 1, 1, 1, raw="A1"),)
    assert literal.spill_reference_tokens == ("A1#",)
    assert literal.tokenization_failed is False
    assert qualified.references == (
        ParsedReference("My Sheet", 2, 2, 2, 2, raw="'My Sheet'!$B$2"),
    )
    assert qualified.spill_reference_tokens == ("'My Sheet'!$B$2#",)
    assert serialized.references == (ParsedReference(None, 1, 1, 1, 1, raw="A1"),)
    assert serialized.spill_reference_tokens == ("_xlfn.ANCHORARRAY",)
    assert string_literal.references == (ParsedReference(None, 1, 1, 1, 1, raw="A1"),)
    assert string_literal.spill_reference_tokens == ("A1#",)


def test_formula_inspection_marks_unsupported_tokenization_failures() -> None:
    malformed = inspect_formula("=SUM(A1#1)")
    range_anchor = inspect_formula("=SUM(A1:B2#)")
    external_anchor = inspect_formula("=SUM([book.xlsx]Sheet1!A1#)")
    implicit_spill = inspect_formula("=SUM(@A1#)")

    assert malformed.tokenization_failed is True
    assert range_anchor.tokenization_failed is True
    assert external_anchor.tokenization_failed is True
    assert implicit_spill.tokenization_failed is True


def test_formula_inspection_traces_direct_implicit_intersection_and_marks_all_forms() -> None:
    literal = inspect_formula("=SUM(@A1:A3)", origin=("Model", "B2"))
    serialized = inspect_formula(
        "=SUM(_xlfn.SINGLE(Inputs!$B$2:$B$4))", origin=("Model", "C3")
    )
    function = inspect_formula("=@INDEX(A1:A3,B1)", origin=("Model", "B2"))
    dynamic_function = inspect_formula("=@OFFSET(A1,1,0)", origin=("Model", "B2"))
    horizontal = inspect_formula("=@A1:C1", origin=("Model", "B2"))
    rectangular = inspect_formula("=@A1:C3", origin=("Model", "B2"))
    outside_range = inspect_formula("=@A1:A3", origin=("Model", "B5"))
    named_range = inspect_formula("=@NamedRange", origin=("Model", "B2"))
    table_current_row = inspect_formula("=[@Amount]", origin=("Model", "B2"))
    string_literal = inspect_formula('=CONCAT("@A1:A3",@A1:A3)', origin=("Model", "B2"))

    assert literal.references == (ParsedReference(None, 1, 2, 1, 2, raw="A1:A3"),)
    assert literal.unresolved_range_tokens == ()
    assert literal.implicit_intersection_tokens == ("@A1:A3",)
    assert serialized.references == (
        ParsedReference("Inputs", 2, 3, 2, 3, raw="Inputs!$B$2:$B$4"),
    )
    assert serialized.implicit_intersection_tokens == ("_xlfn.SINGLE",)
    assert function.references == (
        ParsedReference(None, 1, 1, 1, 3, raw="A1:A3"),
        ParsedReference(None, 2, 1, 2, 1, raw="B1"),
    )
    assert function.implicit_intersection_tokens == ("@INDEX",)
    assert dynamic_function.dynamic_reference_functions == ("OFFSET",)
    assert dynamic_function.implicit_intersection_tokens == ("@OFFSET",)
    assert horizontal.references == (ParsedReference(None, 2, 1, 2, 1, raw="A1:C1"),)
    assert rectangular.references == (ParsedReference(None, 2, 2, 2, 2, raw="A1:C3"),)
    assert outside_range.references == (ParsedReference(None, 1, 1, 1, 3, raw="A1:A3"),)
    assert named_range.unresolved_range_tokens == ("@NamedRange",)
    assert named_range.implicit_intersection_tokens == ("@NamedRange",)
    assert table_current_row.implicit_intersection_tokens == ()
    assert string_literal.references == (ParsedReference(None, 1, 2, 1, 2, raw="A1:A3"),)
    assert string_literal.implicit_intersection_tokens == ("@A1:A3",)


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


def test_named_lambda_definition_detection_requires_one_valid_top_level_lambda() -> None:
    assert lambda_parameter_count("=LAMBDA(value,value+1)") == 1
    assert lambda_parameter_count("=_xlfn.LAMBDA(42)") == 0
    assert (
        lambda_parameter_count(
            "=_xlfn.LAMBDA(_xlpm.temp,_xlpm.temp+Inputs!B2)"
        )
        == 1
    )
    assert lambda_parameter_count("=_xlfn.LAMBDA(_xlop.value,_xlop.value)") == 1
    assert lambda_parameter_count("=LAMBDA(value,value)(1)") is None
    assert lambda_parameter_count("=LAMBDA(A1,A1)") is None
    assert lambda_parameter_count("=SUM(LAMBDA(value,value)(1))") is None


def test_formula_inspection_expands_named_lambda_calls_without_shadowing_locals() -> None:
    named_functions = {
        "tocelsius": (
            ParsedReference("Inputs", 2, 2, 2, 2, raw="ToCelsius"),
        ),
        "unsafelookup": None,
        "f": (
            ParsedReference("ShouldNotAppear", 1, 1, 1, 1, raw="F"),
        ),
    }
    static_call = inspect_formula(
        "=ToCelsius(A2)", named_function_references=named_functions
    )
    unsafe_call = inspect_formula(
        "=UnsafeLookup(A2)", named_function_references=named_functions
    )
    locally_bound_call = inspect_formula(
        "=LET(f,LAMBDA(x,x+Inputs!B2),f(A2))",
        named_function_references=named_functions,
    )
    serialized_lambda = inspect_formula(
        "=_xlfn.LAMBDA(_xlpm.temp,_xlpm.temp+Inputs!B2)(A2)"
    )
    optional_serialized_lambda = inspect_formula(
        "=_xlfn.LAMBDA(_xlop.temp,_xlpm.temp+Inputs!B2)(A2)"
    )
    spaced_serialized_lambda = inspect_formula(
        "=_xlfn.LAMBDA(_xlop.temp, _xlfn.LET(_xlpm.rate, Inputs!B2, "
        "_xlpm.temp*_xlpm.rate))(A2)"
    )
    serialized_let = inspect_formula("=_xlfn.LET(_xlpm.rate,Inputs!B2,_xlpm.rate*2)")

    assert static_call.references == (
        ParsedReference("Inputs", 2, 2, 2, 2, raw="ToCelsius"),
        ParsedReference(None, 1, 2, 1, 2, raw="A2"),
    )
    assert static_call.unresolved_range_tokens == ()
    assert unsafe_call.references == (
        ParsedReference(None, 1, 2, 1, 2, raw="A2"),
    )
    assert unsafe_call.unresolved_range_tokens == ("UnsafeLookup",)
    assert locally_bound_call.references == (
        ParsedReference("Inputs", 2, 2, 2, 2, raw="Inputs!B2"),
        ParsedReference(None, 1, 2, 1, 2, raw="A2"),
    )
    assert locally_bound_call.unresolved_range_tokens == ()
    assert serialized_lambda.references == (
        ParsedReference("Inputs", 2, 2, 2, 2, raw="Inputs!B2"),
        ParsedReference(None, 1, 2, 1, 2, raw="A2"),
    )
    assert serialized_lambda.unresolved_range_tokens == ()
    assert optional_serialized_lambda.references == serialized_lambda.references
    assert optional_serialized_lambda.unresolved_range_tokens == ()
    assert spaced_serialized_lambda.references == serialized_lambda.references
    assert spaced_serialized_lambda.unresolved_range_tokens == ()
    assert serialized_let.references == (
        ParsedReference("Inputs", 2, 2, 2, 2, raw="Inputs!B2"),
    )
    assert serialized_let.unresolved_range_tokens == ()


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
