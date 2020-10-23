

clear
cls
set trace off
local dnum = string(td(`c(current_date)'),"%tdCYND")
local dstr = string(td(`c(current_date)'),"%tdY.N.D")
cd "C:\Users\adw54\Documents\GitHub\cto2file"


********************************************************************************
*
*	Bihar Migrants Survey
*
********************************************************************************

//	EDIT THESE: Get all formdef versions used in the survey
	* Form definitions in folder
	local versions 2004250008 2004262138 2004290133 2004292216 2004302145 2005010209 2005032034 2005050200 2005052256 2005062043
	
	* Relative path to folder and file stems
	local file_path form_definitions
	local file_stem Bihar_migrants_survey_CATI


	*set trace on
	
foreach svyversion in 2005062043 { 
	*local svyversion 2004250008
	*local x english
	local survey Bihar_migrants_survey_CATI
	local title Bihar Migrants Survey
	
	*set trace on
	foreach x in english {
		local cap = strupper(substr("`x'",1,1)) + substr("`x'",2,.)
		cto2docx using "`survey'_`svyversion'.xlsx", 	save("`survey'_`svyversion'_`x'.xlsx")	///
			omitn()	omittypes()	keepother othersuffix(_other) 											///
			dropnames(no_case starttime endtime username caseid deviceid subscriberid simid devicephonenum )			///
			dropgroup(gender_mismatch_check answered_response_group filter_unqualified consent_check) droptypes(calculate)						///
			title(`title' - `cap')	sub(v`svyversion') pagesize(A4) language(`x')							///
			command() maxheading(2) split intl(5) decl(2)									
	}
	local last_version `svyversion'
}
	
	
	! `survey'_`last_version'_english.xlsx

		exit
		doedit C:\Users\adw54\Documents\GitHub\cto2file/cto2file.ado
		copy C:\Users\adw54\Documents\GitHub\cto2file/cto2file.ado "`c(sysdir_personal)'\c\cto2file.ado", replace

