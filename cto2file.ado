*! version 2.0.2  9may2021 Aaron Wolf, aaron.wolf@u.northwestern.edu
cap program drop cto2file
program define cto2file

	version 15

	syntax using/, 	SAve(string) 				///
					[							///
					DEFault(name) 				///
					Language(name) 				///
					DROPGroups(namelist local)	///
					DROPNames(namelist local) 	///
					DROPTypes(string asis)		///
					OMITNames(namelist local) 	///
					OMITTypes(string asis) 		///
					KEEPTypes(string asis)		///
					OTHERSuffix(name)			///
					OTHERNames(namelist local)	///
					KEEPOther 					///
					KEEPRelevance				///
					Title(string) 				///
					SUBtitle(string)			///
					MAXheading(integer 3)		///
					INTlength(integer 9)		///
					DEClength(integer 9)		///
					Pagesize(passthru)			///
					split						///
					novarnum					///
					noheadnum					///
					COmmand(string asis)		///
					]


	clear
	
//	Syntax Checks
	* Ensure that Input document is excel, and output is docx
	if strmatch("`using'","*.xls") != 1 & strmatch("`using'","*.xlsx") != 1 {
		di as error "Using file extension must be .xls or .xlsx"
		exit 198
	}

	if 		strmatch("`save'","*.docx") == 1 	local file "docx"
	else if strmatch("`save'","*.xlsx") == 1 	local file "xlsx"
	else {
		di as error "Saving file extension must be .docx or .xlsx"
		exit 198
	}

	* Ensure Omit, Keep, and Drop types are valid
	foreach type in `droptype' `omittype' `keeptype' {
		cap assert 	inlist(	"`type'","audio","audio audit","barcode",			///
							"begin group","begin repeat","calculate",			///
							"calculate_here","caseid","comments","date",		///
							"datetime","decimal","deviceid","end",				///
							"end group","end repeat","file","geopoint",			///
							"image","integer","note","phonenumber",				///
							"select_multiple","select_one","simserial",			///
							"speed violations audit","speed violations count",	///
							"speed violations list","start","subscriberid",		///
							"text","text audit","video")						///
						| strmatch("`type'","select_one *") == 1				///
						| strmatch("`type'","select_multiple *") == 1
			if _rc != 0 {
				di as error "`type' is not a valid SurveyCTO type"
				exit 198
				}
		}

	* Set defaut langauge to English if default is blank
	if "`default'" == "" local default "english"
	*  Set language equal to the default language if language is blank
	if "`language'" == "" {
		di as result "No language specified. `default' assumed."
		local language "`default'"
	}

	* Generate local macro for types to omit from question numbers
	local omitttypes_default `""begin group" "begin repeat" "comments" "end group" "end repeat" "note" "username" "caseid" "start" "end" "deviceid" "text audit" "audio audit""'
	local omittypes : list omitttypes_default | omittypes
	local omittypes : list omittypes - keeptypes

//	Import Survey sheet and begin cleaning
	* Import Survey file
	qui import excel using "`using'", firstrow sheet("survey") clear
	qui drop if missing(type)

	* Execute optional commands
	local replace `"`command'"'
	foreach x in `command' {
		`x'
	}
	

	* Remove all leading and trailing blanks from string variables
	qui ds, has(type string)
	foreach v in `r(varlist)' {
		qui replace `v' = strtrim(`v')
	}

	* Rename default label and hint (no specified language) with default language
	cap confirm variable label`language', exact
		if _rc {
			cap confirm variable label, exact
			if _rc {
				di as error "There is no column named label:`language' or simply label in survey sheet. Please select a different language."
				error 198				
			}
			else rename label label`language'
		}
	
	cap confirm variable hint`language', exact
	if _rc {
		cap confirm variable hint, exact
		if !_rc rename hint hint`language'
		
	}

	* Keep labels and hints for the language we are exporting in
	keep type name label`language' hint`language' constraint relevance disabled response_note required
	qui gen sort = _n

	* Drop unwanted variables in "drop" macros
	foreach n of local dropnames {
		qui drop if name == "`n'"
	}
	foreach t of local droptypes {
		qui drop if type == "`t'"
	}
	foreach n of local dropgroups {
		qui drop if name == "`n'"
	}

	* Generate Question variable
	qui gen question = 1
		* Omitted types are not questions
		foreach t of local omittypes {
			qui replace question = 0 if type == "`t'"
		}
		* Omitted names are not questions
		foreach n of local omitnames {
			qui replace question = 0 if name == "`n'"
		}
		* "Other" names are not questions
		gen other = 0
		if "`othersuffix'" != "" {
			qui replace question = 0 if strmatch(name,"*`othersuffix'") == 1
			qui replace other = 1	 if strmatch(name,"*`othersuffix'") == 1
		}
		foreach n of local othernames {
			qui replace questions == 0 if name == "`n'"
			qui replace other = 1 if name == "`n'"
		}

	* Generate Question Variable Numbers
	qui bysort question (sort): gen varnumber = "V" + string(_n,"%03.0f") if question == 1

	* Generate Note numbers
	qui gen isnote = type == "note"
	qui bysort isnote (sort): gen notenumber = "__n" + string(_n,"%02.0f") if isnote == 1
	qui replace varnumber = notenumber if question != 1 & isnote == 1


	* Drop "other" variables entirely unless "keepother" is specified
	if "`keepother'" != "keepother" qui keep if other == 0

	* Separate out "select_one" and "select_multiple" variable types
	qui split type, parse(" ") gen(type)
	qui replace type1 = "" if !inlist(type1,"select_one","select_multiple")
	qui replace type2 = "" if !inlist(type1,"select_one","select_multiple")

	* Generate Heading Level and drop
	sort sort
	qui gen heading = sum(inlist(type,"begin group","begin repeat") - inlist(type,"end group","end repeat")) // This calculates the heading level of each begin_group statement
	qui gen group0 = 1
	if "`maxheading'" == "" {
		qui sum heading
		local maxheading `r(max)'		
	}	
	forvalues i = 1/`maxheading' {
		local ph = `i' - 1
		local groups = "`groups' group`ph'"
		qui bysort `groups' (sort): gen group`i' = sum(heading[_n-1] == `ph' & heading > `ph') if heading > `ph'
		qui replace group`i' = 0 if group`i' == .
		qui sum group`i'
		local maxlen = strlen(string(`r(max)'))
	}
	forvalues i = 1/`maxheading' {
		qui tostring group`i', replace usedisplayformat
	}
	drop group0
	qui egen module = concat(group*), punct(".")
	qui replace module = group1 if heading == 1
	qui egen tablename = concat(group*)
		qui replace tablename = "table_" + tablename


	* Make relevance strings more readable
	if "`keeprelevance'" == "" {
		qui replace relevance = subinstr(relevance,"''","BLANK",.)
		qui replace relevance = subinstr(relevance,"{","[",.)
		qui replace relevance = subinstr(relevance,"}","]",.)
		foreach b in "$" "selected(" "')" "'" {
			qui replace relevance = subinstr(relevance,"`b'","",.)
		}
		qui replace relevance = subinstr(relevance,",","=",.)
		qui replace relevance = stritrim(relevance)
		qui replace relevance = strtrim(relevance)
		foreach sign in "=" ">" "<" "!"{
			qui replace relevance = subinstr(relevance," `sign' ","`sign'",.)
			qui replace relevance = subinstr(relevance," `sign'","`sign'",.)
			qui replace relevance = subinstr(relevance,"`sign' ","`sign'",.)
			qui replace relevance = subinstr(relevance,"`sign'"," `sign' ",.)
		}
		qui replace relevance = stritrim(relevance)
		qui replace relevance = strtrim(relevance)
		foreach sign in ">" "<" "!" {
			qui replace relevance = subinstr(relevance,"`sign' =","`sign'=",.)
		}

		forvalues i = 1 / `c(N)' {
			if !inlist(`type[`i']',"end group","end repeat") {
				local name = name[`i']
				if "`varnum'" == "" local var  = varnumber[`i'] + " "
				qui replace relevance = subinstr(relevance,"[`name']","`var'[`name']",.)
			}
		}
	}


	* Replace "$" and "{" or "}" with [ ]
	foreach x in label hint {
		qui cap tostring `x'`language', replace
		qui replace `x'`language' = "" if `x'`language' == "."
		qui replace `x'`language' = subinstr(`x'`language',"$","",.)
		qui replace `x'`language' = subinstr(`x'`language',"{","[",.)
		qui replace `x'`language' = subinstr(`x'`language',"}","]",.)
	}

	* Replace apostrophe's with \'
	qui replace label`language' = subinstr(label`language',"'","`=char(39)'",.)
	qui replace hint`language' = subinstr(hint`language',"'","`=char(39)'",.)
	qui replace label`language' = subinstr(label`language',"`","`=char(39)'",.)
	qui replace hint`language' = subinstr(hint`language',"`","`=char(39)'",.)

	* Preserve line breaks
	qui replace label`language' = subinstr(label`language',"`=char(10)'`=char(10)'","`=char(10)'",.)	// Collapse all spaces to one
	qui replace hint`language' = subinstr(hint`language',"`=char(10)'`=char(10)'","`=char(10)'",.)		// Collapse all spaces to one

	qui count if label`language' != ""
	if `r(N)' > 0 qui split label`language', parse("`=char(10)'") gen(lblsgmt)
	else gen lblsgmt1 = label`language'
	
	qui count if hint`language' != ""
	if `r(N)' > 0 qui split hint`language', parse("`=char(10)'") gen(hintsgmt)
	else gen hintsgmt1 = hint`language'

	* Replace double quote with single quotes
	qui replace label`language' = subinstr(label`language',`"""',"`=char(39)'",.)
	qui replace hint`language' = subinstr(hint`language',`"""',"`=char(39)'",.)


	* Create variable for number of columns before and after decimal (based on constraint)
	quietly {
	cap assert mi(constraint)
	if _rc != 0 { // begin constraint if
		gen constraint_orig = constraint
		replace constraint = "" if !inlist(type,"integer","decimal")
		gen length = strlen(constraint)
		qui sum length
		forvalues i = 1/`r(max)' {
			replace constraint = subinstr(constraint,substr(constraint,`i',1)," ",.) 		///
				if 	!inlist(substr(constraint,`i',1),"0","1","2","3","4","5","6","7") 	& 	///
					!inlist(substr(constraint,`i',1),"8","9"," ",".","-")					//
			}
		replace constraint = strtrim(constraint)
		replace constraint = stritrim(constraint)
		split constraint, parse(" ") gen(con_)
		foreach v of varlist con_* {
			split `v', parse(".") gen(`v'_)
			replace `v'_1 = "" if `v'_1 == "-"
			destring `v'_1, replace
			replace `v'_1 = abs(`v'_1)
			cap replace `v'_2 = "" if `v'_2 == "-"
			cap destring `v'_2, replace
			cap replace `v'_2 = abs(`v'_2)
		}
		cap egen min_predec = rowmin(con_?_1)
			if _rc == 111 gen min_predec = .
		cap egen max_predec = rowmax(con_?_1)
			if _rc == 111 gen max_predec = .
		cap egen min_postdec = rowmin(con_?_2)
			if _rc == 111 gen min_postdec = .
		cap egen max_postdec = rowmax(con_?_2)
			if _rc == 111 gen max_postdec = .

		foreach v of varlist min_predec max_predec min_postdec max_postdec {
			replace `v' = length(string(`v')) if `v' != .
		}
		egen columns_pre = rowmax(min_predec max_predec)
		egen columns_post = rowmax(min_postdec max_postdec)
		replace columns_pre = 5 if inlist(type,"integer","decimal") & missing(columns_pre) & missing(columns_post)
		replace columns_post = 2 if inlist(type,"integer","decimal") & missing(columns_pre) & missing(columns_post)
		drop length con_* min_* max_* constraint
		rename constraint_orig constraint
	}	// End contraint if
	else { // Begin constraing else
		gen columns_pre = 5 if inlist(type,"integer","decimal") 
		gen columns_post = 2 if inlist(type,"integer","decimal")
	} // End constraing else
	} // End quietly
	* Generate row number
	sort sort
	qui gen row = _n
	qui sum row
	local rows = `r(max)'


//	Send choices to mata tables
	tempfile survey
	qui save `survey'

	* Generate tables of choices for each choice-set
	qui import excel using "`using'", firstrow sheet("choices") clear
	qui rename label label`default'
	cap qui replace value = value + "."
		if _rc == 109 {
			qui tostring value, replace
			qui replace value = value + "."
		}
	qui replace label`language' = label`default' if missing(label`language')
	qui replace label`language' = subinstr(label`language',"'","`=char(39)'",.)
	qui replace label`language' = subinstr(label`language',"`","`=char(39)'",.)

	* Replace "$" and "{" or "}" with [ ]
	qui replace label`language' = subinstr(label`language',"$","",.)
	qui replace label`language' = subinstr(label`language',"{","[",.)
	qui replace label`language' = subinstr(label`language',"}","]",.)

	qui replace list_name = strtrim(stritrim(list_name))
	qui levelsof list_name, local(lists)
	foreach l of local lists {
		cap confirm name `l'		// Ensure table name can be real name
			if _rc == 7 {
				di as error "`l' is not a valid table name."
				exit 7
			}
		else local choicetab "width(100%)"
		cap assert !missing("`l'")
		if _rc != 0 {
			di as error "`l' is blank"
			exit
		}
		
		* Pull choice list to mata matrix
		gen keep = list_name == "`l'"						// List indicator
		cap confirm string variable value
			if _rc 	mata: values = st_data(.,("value"))		// Pull real data if values are real
			else 	mata: values = st_sdata(.,("value"))	// Pull string data if values have strings
		mata: labels = st_sdata(.,("label`language'"))		// Pull labels
		mata: keep = st_data(.,("keep"))
		mata: `l' = select((values,labels),keep)			// Keep values and labels from list
		drop keep
	}

	* Generate "Underline" for text variables
	mata: text = " "

	
********************************************************************************
*
*		Write Document
*
********************************************************************************

**** Word Document *************************************************************
if "`file'" == "docx" { 						// Begin putdocx if statement
//	Re-load survey data
	qui use `survey', clear

	* Replace varnumber as name if option novarnum is specified
	qui if "`varnum'" == "novarnum" replace varnumber = name if !inlist(type,"end group","end repeat")
	sort sort

//	Local macros for styles used in putdocx
	* Greys
	local g_1 222222
	local g_2 4a4a4a
	local g_3 978d85
	local g_4 dddddd
	local g_5 f9f9f9
	
	* Colors
	local c_1 	00356b	// 	Yale Blue
	local c_2	286dc0	//	Yale Medium Blue
	local c_3	63aaff	//	Yale Light Blue
	local c_4	5f712d	//	Yale Green
	local c_5	bd5319	//	Yale Orange

	* Create local macros for each heading level and font style for the table
	local nametext 			`"font(Calibri,8,"`g_1'") italic"'
	local relevancetext 	`"font(Calibri,8,"`g_3'") italic"'
	local endtext 			`"font("Calibri Light",8,"`g_4'") bold italic"'
	local labeltext			`"font(Calibri,10)"'
	local hinttext			`"font(Calibri,8,"`g_3'") italic"'
	local choicetab 		"layout(autofitcontents) cellmargin(right,0pt)"

//	Begin document
	putdocx clear
	putdocx begin, `pagesize' `labeltext' footer(main)
	
	* Title and Subtitle
	putdocx paragraph, 	style(Title)
		putdocx text ("`title'")
	putdocx paragraph, 	style(Subtitle)
		putdocx text ("`subtitle'")
	
	* Add page numbers
	putdocx paragraph, tofooter(main) halign(center)
	putdocx pagenumber

//	Loop through all rows in dataset and write contents to file
	forvalues i = 1/`rows' {
		** Group Headings
		if type[`i'] == "begin group" | type[`i'] == "begin repeat" {
			local h = heading[`i']
			if heading[`i'] == 1 local append append
				else local append ""
			if type[`i'] == "begin repeat" local repeat " (Repeat Group)"
				else local repeat ""
			local label = label`language'[`i']
			local module = module[`i']
			local heading = "Heading" + strofreal(heading[`i'])
			if !missing(relevance[`i']) local linebreak linebreak
				else local linebreak ""
				
			if heading[`i'] == 1 putdocx sectionbreak

			putdocx paragraph, style(`heading')
			if "`headnum'" == "" putdocx text ("`module' - `label'`repeat'"), `linebreak'
			else if "`headnum'" == "noheadnum" putdocx text ("`label'`repeat'"), `linebreak'
			putdocx paragraph
			putdocx text (relevance[`i']), `relevancetext'
				
		}
		else if type[`i'] == "end group" |  type[`i'] == "end repeat" {
			qui sum row if name == name[`i']
			local label = label`language'[`r(min)']
			local module = module[`r(min)']
			if type[`i'] == "end repeat" local repeat "Repeat "
			else local repeat ""

			putdocx paragraph,  halign(right)
			if "`headnum'" == "" putdocx text ("	End `repeat'Group: `module' - `label'"),  `endtext'
			else if "`headnum'" == "noheadnum" putdocx text ("	End `repeat'Group: `label'"),  `endtext'
			*putdocx table surveytable(`i',1) = ("	End `repeat'Group: `module' - `label'"), colspan(`columns') `endtext' halign(right)
		}
		
		** Notes and Question Text
		else {
			* Special Option Text
			if !inlist(type[`i'],"begin_group","end_group","begin_repeat","end_repeat") & !inlist(type[`i'],"text","integer","decimal","date","datetime","note","calculate","calculate_here") & !inlist(type1[`i'],"select_one","select_multiple") {
				local special = "[" + type[`i'] + "] "
			}
			else local special ""			
			
			* Label text
			if type[`i'] == "note" {
					putdocx paragraph, halign(left)
					putdocx text (""), linebreak
			}
			else putdocx paragraph, halign(left) indent(hanging,25pt)
			
			putdocx text (cond(type[`i'] != "note",varnumber[`i']+". `special'","") ), `labeltext'
			foreach x of varlist lblsgmt* {
				if !missing(`x'[`i']){
					local segnum = real(subinstr("`x'","lblsgmt","",.))
					local next = `segnum'+1
					cap confirm variable lblsgmt`next'
					if _rc == 0 {
						if !missing(lblsgmt`next'[`i']) local linebreak linebreak(2)
						else if !missing(hintsgmt1[`i']) local linebreak linebreak
						else local linebreak ""
					}
					else if !missing(hintsgmt1[`i']) local linebreak linebreak
					else local linebreak ""
					putdocx text (`x'[`i'] ), `labeltext' `linebreak'
				}
			}
			
			* Hint Text
			foreach x of varlist hintsgmt* {
				if !missing(`x'[`i']) {
					local segnum = real(subinstr("`x'","hintsgmt","",.))
					local next = `segnum'+1
					cap confirm variable hintsgmt`next'
					if _rc == 0 {
						if !missing(hintsgmt`next'[`i']) local linebreak linebreak
						else local linebreak ""
					}
					else local linebreak ""
					putdocx text (`x'[`i']), `hinttext' `linebreak'
				}
			}
		}
		
		** Response Options
		* Calculations
		if inlist(type[`i'],"calculate","calculate_here") {
				putdocx paragraph, halign(left) indent(left,25pt)
				putdocx text ("Calculated Value"), `nametext'	
		}
		
		* Select One/Multiple
		else if type1[`i'] == "select_one" | type1[`i'] == "select_multiple"	{
			if type1[`i'] == "select_multiple" {
				putdocx paragraph, halign(left) indent(left,25pt)
				putdocx text ("Select all that apply:"), `text' italic
			}
			local tname = name[`i']
			local table = type2[`i']
			putdocx table `tname' = mata(`table'), indent(40pt) border(all, nil) `choicetab'
			}
		
		* Text Input
		else if type[`i'] == "text" {
			local tname = name[`i']
			putdocx table `tname' = mata(text), width(2in) indent(40pt) border(all,nil) border(bottom,thick,`g_2')
			}
		
		* Interger Input
		else if type[`i'] == "integer" {
			local tblname = name[`i'] + "_int"
			local columns_pre = columns_pre[`i']
			if `columns_pre' == . local columns_pre = `intlength'
			putdocx table `tblname' = (1,`columns_pre'), border(all,single,`gs_2') border(top,nil) layout(autofitcontents) indent(40pt)
			}
		
		* Decimal Input
		else if type[`i'] == "decimal" {
			local tblname = name[`i'] + "_dec"
			local columns_pre = columns_pre[`i']
			if `columns_pre' == . local columns_pre = `intlength'
			local columns_post = columns_post[`i']
			if `columns_post' == . local columns_post = `declength'
			local totalcolumns_dec = `columns_pre' + `columns_post' + 1
			local decimalpoint = `columns_pre' + 1
			putdocx table `tblname' = (1,`totalcolumns_dec'), border(all,single,`gs_2') border(top,nil) layout(autofitcontents) indent(40pt)
				putdocx table `tblname'(1,`decimalpoint') = ("."), border(top,nil) border(bottom,nil)
			}		

		* Date
		else if type[`i'] == "date" {
			local tname = name[`i']
			putdocx table `tname' = (1,10), border(all,single,`gs_2') border(top,nil) width(0.5in) indent(40pt)
			putdocx table `tname'(1,1) = ("d"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,2) = ("d"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,3) = ("/"), border(top,nil) border(bottom,nil)
			putdocx table `tname'(1,4) = ("m"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,5) = ("m"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,6) = ("/"), border(top,nil) border(bottom,nil)
			putdocx table `tname'(1,7) = ("y"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,8) = ("y"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,9) = ("y"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,10) = ("y"), border(top,nil) font(, , lightgray)
		}
		
		* Date-time
		else if type[`i'] == "datetime" {
			local tname = name[`i']
			putdocx table `tname' = (1,19), border(all,single,`gs_2') border(top,nil) width(0.5in) indent(40pt)
			putdocx table `tname'(1,1) = ("d"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,2) = ("d"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,3) = ("/"), border(top,nil) border(bottom,nil)
			putdocx table `tname'(1,4) = ("m"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,5) = ("m"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,6) = ("/"), border(top,nil) border(bottom,nil)
			putdocx table `tname'(1,7) = ("y"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,8) = ("y"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,9) = ("y"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,10) = ("y"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,11) = ("-"), border(top,nil) border(bottom,nil)
			putdocx table `tname'(1,12) = ("h"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,13) = ("h"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,14) = (":"), border(top,nil) border(bottom,nil)
			putdocx table `tname'(1,15) = ("m"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,16) = ("m"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,17) = (":"), border(top,nil) border(bottom,nil)
			putdocx table `tname'(1,18) = ("s"), border(top,nil) font(, , lightgray)
			putdocx table `tname'(1,19) = ("s"), border(top,nil) font(, , lightgray)
		}

		
		
	} // End row loop

	putdocx save "`save'", replace
	di as text "Document saved as " as result `"`save'"'
	sort sort
	
}													// End putdocx if statement
	
	
**** Excel Sheet ***************************************************************
if "`file'" == "xlsx" { 						// Begin putexcel if statement

//	Re-load survey data
qui {
	gen lsort = _n
	gen type2 = list_name
	gen choicelabel = label`language'
	keep type2 value choicelabel lsort
	drop if value == "."
	set obs `=_N+1'
	replace value = "." in `=_N'
	tempfile choices
	save `choices'
	
	qui use `survey', clear
	joinby type2 using `choices'

	* Replace varnumber as name if option novarnum is specified
	qui if "`varnum'" == "novarnum" replace varnumber = name if !inlist(type,"begin group","begin repeat","end group","end repeat")
	sort sort lsort	
}
//	Local macros for styles used in putexcel
	* Greys
	local g_1 034 034 034
	local g_2 074 074 074
	local g_3 151 141 133
	local g_4 221 221 221
	local g_5 249 249 249
	
	* Colors
	local c_1 	000 053 107	// 	Yale Blue
	local c_2	040 109 192	//	Yale Medium Blue
	local c_3	099 170 255	//	Yale Light Blue
	local c_4	095 113 045	//	Yale Green
	local c_5	189 083 025	//	Yale Orange

	* Create local macros for each heading level and font style for the table
	local titletext			`"font("Calibri Light",28,"white")"'
	local subtitletext		`"font(Calibri,11,"white")"'
	local nametext 			`"font(Calibri,8,"`g_1'") italic"'
	local relevancetext 	`"font(Calibri,8,"`g_3'") italic"'
	local endtext 			`"font("Calibri Light",8,"`g_2'") bold italic border(top,medium,"`c_1'")"'
	local labeltext			`"font(Calibri,10)"'
	local hinttext			`"font(Calibri,8,"`g_3'") italic"'
	local choicetab 		"layout(autofitcontents) cellmargin(right,0pt)"
	local h1 				`"font("Calibri Light",16,white) fpattern(solid,"`c_2'") border(bottom,medium,"`c_1'") "'
	local h2 				`"font("Calibri Light",13,"`c_2'") fpattern(solid,"`g_4'")"'
	local h3 				`"font("Calibri Light",12,"`c_2'") fpattern(solid,"`g_5'")"'
	local h4 				`"font("Calibri Light",11,"`c_4'") bold"'
	local h5 				`"font("Calibri Light",11,"`c_4'")"'
	local h6 				`"font("Calibri Light",11,"`c_5'") italic"'
	local h6 				`"font("Calibri Light",10,"`c_5'") italic"'
	local h7 				`"font("Calibri Light",9,"`c_5'") italic"'
	local h8 				`"font("Calibri Light",8,"`c_5'") italic"'
	local h9 				`"font("Calibri Light",7,"`c_5'") italic"'

	
	
//	Write mock dataset using stata data
qui {
	foreach x in A B C D {
		cap drop `x'
		gen `x' = ""
	}
	bysort sort (lsort): replace A = varnumber if _n == 1 & type != "note"
	bysort sort (lsort): replace B = label`language' + cond(!mi(hint`language'),"`=char(10)'" + hint`language',"") if _n == 1
	replace B = cond("`headnum'"=="",module + " - ","") + label`language' + cond(type=="begin repeat"," (Repeat Group)","") if inlist(type,"begin group","begin repeat")
	bysort name (sort lsort): replace B = "End " + cond(type=="end repeat","Repeat ","") + "Group: " + cond("`headnum'"=="",module[1] + " - ","") + label`language'[1] if inlist(type,"end group","end repeat")
	sort sort lsort
	replace D = "Calculated Value" 							if inlist(type,"calculate","calculate_here")
	replace D = "___________________" 						if type == "text"
	replace D = "|_|_|_|_|_|" 								if type == "integer"
	replace D = "|_|_|_|_|_|.|_|_|" 						if type == "decimal"
	replace D = "|_|_|/|_|_|/|_|_|_|_|"						if type == "date"
	replace D = "|_|_|/|_|_|/|_|_|_|_| - |_|_|:|_|_|:|_|_|" if type == "datetime"
	replace C = value 										if inlist(type1,"select_one","select_multiple")
	replace D = choicelabel									if inlist(type1,"select_one","select_multiple")

	sort sort lsort
	gen n = _n
	set obs `=_N+1'
	replace A = "`title'" in `=_N'
	replace n = -2 in `=_N'
	set obs `=_N+1'
	replace A = "`subtitle'" in `=_N'
	replace n = -1 in `=_N'
	sort n
	gen rownum = _n
}
	qui export excel A B C D using "`save'", replace
	
//	Format Rows
	qui putexcel clear
	qui putexcel set "`save'", open modify
	
	* Start with Plain white, no borders
	qui putexcel B1:D`c(N)' , overwritefmt fpattern(solid,white) left top
	
	* Title/Subtitle
	qui putexcel (A1:D1) , overwritefmt merge fpattern(solid,"`c_1'") `titletext' txtindent(5)
	qui putexcel (A2:D2) , overwritefmt merge fpattern(solid,"`c_1'") `subtitletext' txtindent(5)
	
	* Left sidebar
	qui putexcel A3:A`c(N)' , overwritefmt fpattern(solid,"`c_3'") font(Calibri,12,white) left top
	
	* Headings
	qui levelsof name if inlist(type,"begin group","begin repeat"), local(names)
	foreach name of local names {		
		qui sum rownum if name == "`name'"
		local r = `r(min)'
		local s = `r(max)'
		local heading = "h" + string(heading[`r'])
		if heading[`r'] == 1 qui putexcel A`r', 		overwritefmt  	   ``heading'' left top
		qui putexcel B`r':D`r', overwritefmt merge ``heading'' left top
		qui putexcel B`s':D`s', overwritefmt merge  `endtext' right top
		if heading[`r'] == 1  qui putexcel A`s', 		overwritefmt 		`endtext' right top
	}
	
	* Top border between questions
	qui levelsof name if !inlist(type,"begin group","begin repeat","end group","end repeat"), local(names)
	foreach name of local names {		
		qui sum rownum if name == "`name'"
		local r = `r(min)'
		local s = `r(max)'
		qui putexcel B`r':D`r', overwritefmt border(top,thin,"`c_2'") left top
		qui putexcel A`r', overwritefmt fpattern(solid,"`c_3'") font(Calibri,12,white) left top border(top,thin,"`c_2'")
	}
	
	* Merge cells for select_one/select_multiple questions
	qui levelsof name if inlist(type1,"select_one","select_multiple"), local(names)
	foreach name of local names {		
		qui sum rownum if name == "`name'"
		local r = `r(min)'
		local s = `r(max)'
		qui putexcel B`r':B`s', overwritefmt merge border(top,thin,"`c_2'") left top
	}
	
// 	* Merge C and D for non select variables
// 	qui levelsof name if !inlist(type1,"select_one","select_multiple") & !inlist(type,"begin group","begin repeat","end group","end repeat"), local(names)
// 	foreach name of local names {		
// 		qui sum rownum if name == "`name'"
// 		local r = `r(min)'
// 		local s = `r(max)'
// 		*qui putexcel C`r':D`r', overwritefmt merge border(top,thin,"`c_2'") left vcenter
// 	}
//	
	* Merge notes across
	qui levelsof name if inlist(type,"note"), local(names)
	foreach name of local names {		
		qui sum rownum if name == "`name'"
		local r = `r(min)'
		local s = `r(max)'
		qui putexcel B`r':D`r', overwritefmt merge border(top,thin,"`c_2'") left top
	}
	

	
	putexcel save
	
	
} 												// End putexcel if statement
	
end
