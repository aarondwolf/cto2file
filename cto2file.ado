*! version 1.0.1  19mar2018 Aaron Wolf, awolf@pih.org
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
					scheme(name)				///
					Pagesize(passthru)			///
					split						///
					novarnum					///
					COmmand(string asis)		///
					]
					

	clear				
	* Ensure that Input document is excel, and output is either pdf or docx
	
	if strmatch("`using'","*.xls") != 1 & strmatch("`using'","*.xlsx") != 1 {
		di as error "Using file extension must be .xls or .xlsx"
		exit 198
	}
	
	if strmatch("`save'","*.pdf") == 1 			local doc "pdf"
	else if strmatch("`save'","*.docx") == 1 	local doc "docx"
	else {
		di as error "Saving file extension must be .pdf or .docx"
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
	
	* Confirm scheme is in our list
	if "`scheme'" == "" local scheme "blue"
	cap assert inlist("`scheme'","blue","red","orange","purple","green","colorful","drab")
		if _rc == 9 {
			di as error "Invalid color scheme"
			exit 198
			}
	confirm names `scheme'

	* Set defaut langauge to English if default is blank
	if "`default'" == "" local default "english"
	*  Set language equal to the default language if language is blank
	if "`language'" == "" local language "`default'"
	
	* Generate local macro for types to omit from question numbers
	local omitttypes_default `""begin group" "begin repeat" "comments" "end group" "end repeat" "note" "username" "caseid" "start" "end" "deviceid" "text audit" "audio audit""'
	local omittypes : list omitttypes_default | omittypes
	local omittypes : list omittypes - keeptypes
	
	
	* Import Survey file
	qui import excel using "`using'", firstrow sheet("survey") clear
	qui drop if missing(type)
	
	* Execute optional "replace" commands
	local replace `"`command'"'
	foreach x in `command' {
		`x'
	}
	
	* Remove all leading and trailing blanks from string variables
	qui ds, has(type string)
	foreach v in `r(varlist)' {
		qui replace `v' = strtrim(`v')
	}
	
	* Rename default label and hint with default language
	rename label label`default'
	rename hint hint`default'
	
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
	qui gen heading = sum(inlist(type,"begin group","begin repeat") - inlist(type,"end group","end repeat"))
	qui gen group0 = 1
	forvalues i = 1/`maxheading' {
		local ph = `i' - 1
		local groups = "`groups' group`ph'"
		qui bysort `groups' (sort): gen group`i' = sum(heading[_n-1] == `ph' & heading > `ph') if heading > `ph'
		qui replace group`i' = 0 if group`i' == .
		qui sum group`i'
		local maxlen = strlen(string(`r(max)'))
		*format %0`maxlen'.0f group`i'
	}
	forvalues i = 1/`maxheading' {
		qui tostring group`i', replace usedisplayformat
	}
	drop group0
	qui egen module = concat(group*), punct(".")
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
	qui replace hint`language' = subinstr(hint`language',"`=char(10)'`=char(10)'","`=char(10)'",.)	// Collapse all spaces to one
	
	qui split label`language', parse("`=char(10)'") gen(lblsgmt)
	qui split hint`language', parse("`=char(10)'") gen(hintsgmt)
	
	* Replace double quote with single quotes
	qui replace label`language' = subinstr(label`language',`"""',"`=char(39)'",.)
	qui replace hint`language' = subinstr(hint`language',`"""',"`=char(39)'",.)

	
	* Create variable for number of columns before and after decimal (based on constraint)
	quietly {
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
		replace columns_pre = 2 if inlist(type,"integer","decimal") & missing(columns_pre) & missing(columns_post)
		drop length con_* min_* max_* constraint
		rename constraint_orig constraint
	}
	* Generate row number
	sort sort
	qui gen row = _n
	qui sum row
	local rows = `r(max)'
	
	* Create color palettes
	if "`scheme'" == "blue" {
		local c_1 	031926
		local c_2	032438
		local c_3	0c4363
		local c_4	0e557f
		local c_5	277cad
		}
	else if "`scheme'" == "green" {
		local c_1 	2f3627
		local c_2	5c723f
		local c_3	768c59
		local c_4	73914d
		local c_5	5c664f
		}
	else if "`scheme'" == "purple" {
		local c_1 	32021f
		local c_2	492d3e
		local c_3	6d5865
		local c_4	8c6e80
		local c_5	a87594
		}
	else if "`scheme'" == "red" {
		local c_1 	3d1515
		local c_2	7a2d2d
		local c_3	7f3d3d
		local c_4	935e5e
		local c_5	8a6161
		}
	else if "`scheme'" == "orange" {
		local c_1 	3d2015
		local c_2	68270d
		local c_3	a32e00
		local c_4	d36135
		local c_5	f2c1ae
		}
	else if "`scheme'" == "drab" {
		local c_1 	363020
		local c_2	605b4e
		local c_3	a39265
		local c_4	746f63
		local c_5	c6bda5
		}
	else if "`scheme'" == "colorful" {
		local c_1 	3a2e39
		local c_2	1e555c
		local c_3	f15152
		local c_4	edb183
		local c_5	b29e96
		}
	else {							// Default to blues
		local c_1 	031926
		local c_2	032438
		local c_3	0c4363
		local c_4	0e557f
		local c_5	277cad
		}
		
	* Local Macros for Gray Colors
	local g_1 3f3c3c
	local g_2 585453
	local g_3 706d6c
	local g_4 a7a3a2
	local g_5 ccc7c5

	
	* Create local macros for each heading level and font style for the table
	local titletext			`"font("Calibri Light",28,"`c_1'")"'
	local subtitletext		`"font(Calibri,11,"`c_3'")"'
	local nametext 			`"font(Calibri,8,"`g_1'") italic"'
	local relevancetext 	`"font(Calibri,8,"`g_3'") italic"'
	local endtext 			`"font("Calibri Light",8,"`g_4'") bold italic"'
	local labeltext			`"font(Calibri,10)"'
	local hinttext			`"font(Calibri,8,"`g_3'") italic"'
	
	local h1 `"font("Calibri Light",16,"`c_2'")"'
	local h2 `"font("Calibri Light",13,"`c_3'")"'
	local h3 `"font("Calibri Light",12,"`c_4'")"'
	local h4 `"font("Calibri Light",11,"`c_5'") bold"'
	local h5 `"font("Calibri Light",11,"`c_5'")"'
	local h6 `"font("Calibri Light",11,"`c_5'") italic"'
	local h6 `"font("Calibri Light",10,"`c_5'") italic"'
	local h7 `"font("Calibri Light",9,"`c_5'") italic"'
	local h8 `"font("Calibri Light",8,"`c_5'") italic"'
	local h9 `"font("Calibri Light",7,"`c_5'") italic"'
	
	* Column Macros
	local columns 10
	local lspan 5
	local rspan 4
	
	local nspan = `columns' - 1

	* Begin document
	put`doc' begin, `pagesize' `labeltext'
	put`doc' paragraph, 	`titletext'
		put`doc' text ("`title'")
	put`doc' paragraph, 	`subtitletext'
		put`doc' text ("`subtitle'")
	
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
		if "`doc'" == "docx" local choicetab "layout(autofitcontents) cellmargin(right,0pt)"
		else local choicetab "width(100%)"
		cap assert !missing("`l'")
		if _rc != 0 {
			di as error "`l' is blank"
			exit
		}
		qui put`doc' table `l' = data(value label`language') if list_name == "`l'", border(all,nil) memtable `choicetab'
	}
	
	if "`doc'" == "docx" local layout "layout(autofitcontents)"
	else local layout "width(100%)"
	
	* Generate "Underline" for text variables
	putdocx table text_inside = (1,1), border(all,nil) border(bottom,thick,`g_2') layout(autofitwindow) memtable
	putdocx table text = (1,1), border(all,nil) layout(autofitwindow) memtable cellmargin(bottom,4pt) cellmargin(top,8pt)
		putdocx table text(1,1) = table(text_inside)


	/* Generate "Underline" for text variables --> NEED TO WORK ON THIS.... NOT MAKING THE BOXES RIGHT
	put`doc' table text_inside = (1,1), border(all,nil) border(bottom,single,`g_2') `choicetab' memtable
	if "`doc'" == "docx" local cellmargin "cellmargin(bottom,4pt) cellmargin(top,8pt)"
	put`doc' table text = (1,1), border(all,nil) memtable `choicetab' `cellmargin'
		put`doc' table text(1,1) = table(text_inside)
	*/
	
	
	* Re-load survey data
	qui use `survey', clear	
	
	* Replace varnumber as name if option novarnum is specified
	qui if "`varnum'" == "novarnum" replace varnumber = name if !inlist(type,"end group","end repeat")
	
	
	
	sort sort
	* Create Main Table
	putdocx table surveytable = (`rows',`columns'), border(all,single,`g_5')
	forvalues i = 1/`rows' {
		if "`split'" == "" putdocx table surveytable(`i',.), nosplit
		if type[`i'] == "begin group" | type[`i'] == "begin repeat" {
			local h = heading[`i']
			if heading[`i'] == 1 local append append
			else local append ""
			if type[`i'] == "begin repeat" local repeat " (Repeat Group)"
			else local repeat ""
			local label = label`language'[`i']
			local module = module[`i']
			local heading = "h" + strofreal(heading[`i'])
			if !missing(relevance[`i']) local linebreak linebreak
			else local linebreak ""
			
			if heading[`i'] == 1 putdocx table surveytable(`i',1) = (""), `nametext' border(start,nil) border(end,nil) linebreak
			putdocx table surveytable(`i',1) = ("`module' - `label'`repeat'"), colspan(`columns') ``heading'' `append' `linebreak'
			if !missing(relevance[`i']) putdocx table surveytable(`i',1) = (relevance[`i']), append `relevancetext'
			
			}
		else if type[`i'] == "end group" |  type[`i'] == "end repeat" {
			qui sum row if name == name[`i']
			local label = label`language'[`r(min)']
			local module = module[`r(min)']
			if type[`i'] == "end repeat" local repeat "Repeat "
			else local repeat ""
			
			putdocx table surveytable(`i',1) = ("	End `repeat'Group: `module' - `label'"), colspan(`columns') `endtext' halign(right)
		
		}
		else if type[`i'] == "note" {
			putdocx table surveytable(`i',1) = (varnumber[`i']), font("",10)
			if "`varnum'" != "novarnum"	{
				putdocx table surveytable(`i',2) = (name[`i']), colspan(`nspan') `nametext' linebreak
				if !missing(relevance[`i']) 		putdocx table surveytable(`i',2) = (relevance[`i']), `relevancetext' append linebreak
				}
			else {
				putdocx table surveytable(`i',2) = (""), colspan(`nspan') `nametext'
				if !missing(relevance[`i']) 		putdocx table surveytable(`i',2) = (relevance[`i']), `relevancetext' append linebreak
			}
	
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
				putdocx table surveytable(`i',2) = (`x'[`i']), `labeltext' append `linebreak'
				}
			}
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
				putdocx table surveytable(`i',2) = (`x'[`i']), `hinttext' append `linebreak'
				}
			}			
		}
		else {
			putdocx table surveytable(`i',1) = (varnumber[`i']), font("",10)
				if !missing(relevance[`i']) | !missing(lblsgmt1[`i']) | !missing(hintsgmt1[`i']) local linebreak linebreak
				else local linebreak ""
			if "`varnum'" != "novarnum"	{
				putdocx table surveytable(`i',2) = (name[`i']), `nametext' colspan(`lspan') `linebreak'
				if !missing(lblsgmt1[`i']) | !missing(hintsgmt1[`i']) local linebreak linebreak
				else local linebreak ""
				}
			else {
				putdocx table surveytable(`i',2) = (""), `nametext' colspan(`lspan')
				if !missing(lblsgmt1[`i']) | !missing(hintsgmt1[`i']) local linebreak linebreak
				else local linebreak ""
			}
				
			if !missing(relevance[`i']) putdocx table surveytable(`i',2) = (relevance[`i']), `relevancetext' append `linebreak'
			
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
				putdocx table surveytable(`i',2) = (`x'[`i']), `labeltext' append `linebreak'
				}
			}
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
					putdocx table surveytable(`i',2) = (`x'[`i']), `hinttext' append `linebreak'
				}
			}			
			
			if type[`i'] == "calculate" 		putdocx table surveytable(`i',3) = ("Calculated Value"), `nametext' colspan(`rspan')
			
			else if type1[`i'] == "select_one" | type1[`i'] == "select_multiple"	{
				local table = type2[`i']
				if type1[`i'] == "select_one" local select "Select One:"
				else local select "Select All That Apply:"
				
				putdocx table surveytable(`i',3) = ("`select'"), `text' italic colspan(`rspan')
				putdocx table surveytable(`i',3) = table(`table'), append
				}
				
			else if type[`i'] == "text" {
				putdocx table surveytable(`i',3) = table(text), colspan(`rspan') valign(bottom)
				putdocx table surveytable(`i',3), append
				}
				
			else if type[`i'] == "integer" {
				local tblname = name[`i'] + "_int"
				local columns_pre = columns_pre[`i']
				if `columns_pre' == . local columns_pre = `intlength'
				putdocx table `tblname' = (1,`columns_pre'), border(all,single,`gs_2') border(top,nil) memtable layout(autofitcontents)
				putdocx table surveytable(`i',3) = table(`tblname'), colspan(`rspan') valign(center) halign(center)
				}
				
			else if type[`i'] == "decimal" {
				local tblname = name[`i'] + "_dec"
				local columns_pre = columns_pre[`i']
				if `columns_pre' == . local columns_pre = `intlength'
				local columns_post = columns_post[`i']
				if `columns_post' == . local columns_post = `declength'
				local totalcolumns_dec = `columns_pre' + `columns_post' + 1
				local decimalpoint = `columns_pre' + 1
				putdocx table `tblname' = (1,`totalcolumns_dec'), border(all,single,`gs_2') border(top,nil) memtable layout(autofitcontents)
					putdocx table `tblname'(1,`decimalpoint') = ("."), border(top,nil) border(bottom,nil)
				putdocx table surveytable(`i',3) = table(`tblname'), colspan(`rspan') valign(center) halign(center)
				}

			else putdocx table surveytable(`i',3), colspan(`rspan')
		}
	
	}
	
	
	
	put`doc' save "`save'", replace
	di as text "Document saved as " as result `"`save'"'
	*
	sort sort
end
