program define surveyctocompare
    version 16.0
    syntax , form1(string) form2(string) output(string) [NOFORMAT]

    // -------------------------------------------------------------------------
    // surveyctocompare.ado
    // Author: Kelly Montaño (with ChatGPT assistance)
    // Description: Compare two SurveyCTO forms (survey and choices sheets)
    //              and generate a formatted Excel dashboard of differences.
    // -------------------------------------------------------------------------

    // Validate and safely handle input paths (including those with spaces)
	local file_v1 = `"`form1'"'
	local file_v2 = `"`form2'"'

	display as text "Checking form1 path: `file_v1'"
	display as text "Checking form2 path: `file_v2'"

	capture confirm file `"`file_v1'"'
	if _rc {
		di as error "File `file_v1' not found. Please check the path."
		exit 601
	}
	capture confirm file `"`file_v2'"'
	if _rc {
		di as error "File `file_v2' not found. Please check the path."
		exit 601
	}


    local file_v1 = `"`form1'"'
    local file_v2 = `"`form2'"'
    local out = `"`output'"'

    tempfile survey1 survey2 variables_changes labels_update choices_update choices_update_final ///
             survey_lists choices_ids affected_variables_map ///
             head_v1 head_v2 header_summary t_variables t_labels t_choices t_all

    foreach f in 1 2 {
        if "`f'" == "1" local file = "`file_v1'"
        if "`f'" == "2" local file = "`file_v2'"
        import excel using "`file'", sheet("survey") firstrow clear
        keep name label
        drop if missing(name)
        duplicates drop name, force
        rename label label_f`f'
        gen in_f`f' = 1
        save `survey`f''
    }

    use `survey1', clear
    merge 1:1 name using `survey2', keep(master using) nogen
    gen update_type = ""
    replace update_type = "Removed in new version" if in_f1 == 1 & in_f2 == .
    replace update_type = "Added in new version" if in_f1 == . & in_f2 == 1
    gen variable_label = label_f1
    replace variable_label = label_f2 if missing(variable_label)
    keep if update_type != ""
    rename name updated_variable_name
    keep updated_variable_name variable_label update_type
    save `variables_changes'

    use `survey1', clear
    merge 1:1 name using `survey2', keep(match) nogen
    keep if label_f1 != label_f2
    rename name updated_variable_name
    gen format_change_only = regexm(lower(label_f1 + label_f2), "<b>|<br>|<center|</b>|</center|</br>")
    gen diff_chars = abs(length(label_f1) - length(label_f2))
    gen change_nature = "Significant"
    replace change_nature = "Minor (≤5 characters)" if diff_chars <= 5
    replace change_nature = "Format only" if format_change_only == 1
    replace change_nature = "Format + Minor" if format_change_only == 1 & diff_chars <= 5
    drop format_change_only diff_chars
    order updated_variable_name change_nature label_f1 label_f2
    save `labels_update'

    foreach f in 1 2 {
        if "`f'" == "1" local file = "`file_v1'"
        if "`f'" == "2" local file = "`file_v2'"
        import excel using "`file'", sheet("choices") firstrow clear
        keep list_name value label
        drop if missing(list_name) | missing(value)
        replace list_name = strtrim(lower(list_name))
        replace value = strtrim(lower(value))
        rename label label_f`f'
        gen exists_f`f' = 1
        tempfile c`f'
        save `c`f''
    }

    use `c1', clear
    merge 1:1 list_name value using `c2', keep(match master using) nogen keepusing(label_f2 exists_f2)
    replace exists_f1 = 0 if missing(exists_f1)
    replace exists_f2 = 0 if missing(exists_f2)
    gen change_type = ""
    replace change_type = "Removed" if exists_f1 == 1 & exists_f2 == 0
    replace change_type = "Added" if exists_f1 == 0 & exists_f2 == 1
    replace change_type = "Updated" if exists_f1 == 1 & exists_f2 == 1 & label_f1 != label_f2
    keep if change_type != ""
    save `choices_update'

    import excel using "`file_v2'", sheet("survey") firstrow clear
    keep name type
    drop if missing(type)
    gen list_name = word(strlower(strtrim(type)), 2)
    replace list_name = strtrim(list_name)
    keep if list_name != ""
    rename name affected_variable
    save `survey_lists'

    use `choices_update', clear
    keep list_name
    duplicates drop
    save `choices_ids'

    use `choices_ids', clear
    merge 1:m list_name using `survey_lists', keep(match) nogen
    bysort list_name (affected_variable): gen order = _n
    reshape wide affected_variable, i(list_name) j(order)

    ds affected_variable*
    local avars `r(varlist)'
    if "`avars'" != "" {
        egen affected_variables = concat(`avars'), punct(", ")
        replace affected_variables = regexr(affected_variables, ",( *,)+", ",")
        replace affected_variables = regexr(affected_variables, "^,+", "")
        replace affected_variables = regexr(affected_variables, ",+$", "")
        replace affected_variables = trim(affected_variables)
        foreach var of local avars {
            capture drop `var'
        }
    }
    else {
        gen affected_variables = ""
    }

    save `affected_variables_map'
    use `choices_update', clear
    merge m:1 list_name using `affected_variables_map', keep(master match) nogen
    save `choices_update_final'

    import excel using "`file_v1'", sheet("settings") cellrange(B1:C2) firstrow clear
    gen tag = "Old form"
    save `head_v1'
    import excel using "`file_v2'", sheet("settings") cellrange(B1:C2) firstrow clear
    gen tag = "Updated form"
    save `head_v2'

    use `head_v1', clear
    append using `head_v2'
    order tag version form_id
    label variable tag " "
    label variable version "Version"
    label variable form_id "Form ID"
    save `header_summary'

	foreach t in variables labels choices {
		if "`t'" == "variables" use `variables_changes', clear
		if "`t'" == "labels"    use `labels_update', clear
		if "`t'" == "choices"   use `choices_update_final', clear

		gen category = proper("`t'")

		if "`t'" == "labels" {
			capture drop type
			gen type = change_nature
		}
		else if "`t'" == "variables" {
			capture drop type
			gen type = update_type
		}
		else if "`t'" == "choices" {
			capture drop type
			gen type = change_type
		}


        gen n = 1
        collapse (sum) num_changes = n, by(category type)
        tempfile t_`t'
        save `t_`t''
    }

    use `t_variables', clear
	
	capture confirm file `t_variables'
	if _rc {
		di as error "No variable updates detected. Skipping summary export."
		exit 602
	}

    append using `t_labels'
    append using `t_choices'
    save `t_all'

    use `header_summary', clear
    export excel using "`out'", sheet("Summary") cell("A3") firstrow(varlabels) replace
    use `t_all', clear
    label variable category "Category"
    label variable type "Type of modification"
    label variable num_changes "Number of changes"
    export excel using "`out'", sheet("Summary") cell("A8") firstrow(varlabels) sheetmodify
    putexcel set "`out'", sheet("Summary") modify
    putexcel A1 = ("Form updates summary")

    use `variables_changes', clear
	label variable updated_variable_name "Updated variable name"
	label variable update_type "Type of update"
	label variable variable_label "Variable label (new)"
    export excel using "`out'", sheet("Variable Updates") firstrow(varlabels) sheetmodify
	
    use `labels_update', clear
	keep updated_variable_name label_f1 label_f2 change_nature
	label variable updated_variable_name "Updated variable name"
	label variable label_f1 "Old label"
	label variable label_f2 "New label"
	label variable change_nature "Nature of change"
    export excel using "`out'", sheet("Label Updates") firstrow(varlabels) sheetmodify
	
    use `choices_update_final', clear
	keep list_name change_type value label_f1 label_f2 affected_variables
	label variable list_name "List name"
	label variable value "Value"
	label variable label_f1 "Old label"
	label variable label_f2 "New label"
	label variable change_type "Type of change"
	label variable affected_variables "Affected variables"
    export excel using "`out'", sheet("Choice Updates") firstrow(varlabels) sheetmodify

if ("`noformat'" == "") {
    capture python query
    if _rc == 0 {
        capture file close pyf
			file open pyf using "format_excel.py", write replace
			file write pyf `"import openpyxl"' _n
			file write pyf `"from openpyxl.styles import Font, Alignment, Border, Side"' _n
			file write pyf `"from openpyxl.utils import get_column_letter"' _n
			file write pyf `"file_path = r'`out''"' _n
			file write pyf `"wb = openpyxl.load_workbook(file_path)"' _n
			file write pyf `"bold_font = Font(bold=True, name='Arial', size=11)"' _n
			file write pyf `"regular_font = Font(name='Arial', size=11)"' _n
			file write pyf `"title_font = Font(bold=True, name='Arial', size=14)"' _n
			file write pyf `"centered = Alignment(horizontal='center', vertical='center')"' _n
			file write pyf `"wrap = Alignment(wrap_text=True, vertical='top')"' _n
			file write pyf `"thin_border = Border(bottom=Side(style='thin'))"' _n
			file write pyf `"special_widths = {"' _n
			file write pyf `"    'Variable Updates': {'Variable label (new)': 103},"' _n
			file write pyf `"    'Label Updates': {'Old label': 50, 'New label': 50},"' _n
			file write pyf `"    'Choice Updates': {'Old label': 30.7, 'New label': 30.7, 'Affected variables': 75}"' _n
			file write pyf `"}"' _n
			file write pyf `"for sheet in wb.sheetnames:"' _n
			file write pyf `"    ws = wb[sheet]"' _n
			file write pyf `"    if sheet == 'Summary':"' _n
			file write pyf `"        ws.merge_cells('A1:C1')"' _n
			file write pyf `"        title_cell = ws['A1']"' _n
			file write pyf `"        title_cell.font = title_font"' _n
			file write pyf `"        title_cell.alignment = centered"' _n
			file write pyf `"        for row in [3, 8]:"' _n
			file write pyf `"            for col in range(1, ws.max_column + 1):"' _n
			file write pyf `"                cell = ws.cell(row=row, column=col)"' _n
			file write pyf `"                if cell.value:"' _n
			file write pyf `"                    cell.font = bold_font"' _n
			file write pyf `"                    cell.border = thin_border"' _n
			file write pyf `"                    cell.alignment = centered"' _n
			file write pyf `"    sheet_widths = special_widths.get(sheet, {})"' _n
			file write pyf `"    for col in range(1, ws.max_column + 1):"' _n
			file write pyf `"        col_letter = get_column_letter(col)"' _n
			file write pyf `"        header_cell = ws.cell(row=1, column=col)"' _n
			file write pyf `"        col_name = str(header_cell.value).strip() if header_cell.value else ''"' _n
			file write pyf `"        header_cell.font = bold_font"' _n
			file write pyf `"        header_cell.alignment = centered"' _n
			file write pyf `"        header_cell.border = thin_border"' _n
			file write pyf `"        max_len = len(col_name)"' _n
			file write pyf `"        for row in ws.iter_rows(min_row=2, min_col=col, max_col=col):"' _n
			file write pyf `"            for cell in row:"' _n
			file write pyf `"                cell.font = regular_font"' _n
			file write pyf `"                if col_name in sheet_widths:"' _n
			file write pyf `"                    cell.alignment = wrap"' _n
			file write pyf `"                if cell.value:"' _n
			file write pyf `"                    max_len = max(max_len, len(str(cell.value)))"' _n
			file write pyf `"        if col_name in sheet_widths:"' _n
			file write pyf `"            ws.column_dimensions[col_letter].width = sheet_widths[col_name]"' _n
			file write pyf `"        else:"' _n
			file write pyf `"            ws.column_dimensions[col_letter].width = min(max_len + 2, 50)"' _n
			file write pyf `"wb.save(file_path)"' _n
			file close pyf
			shell python "format_excel.py"

    }
    else {
        di as error "Python is not installed or not available. Run with , noformat to skip formatting."
    }
}


    display as result "Dashboard successfully created: `out'"
end

