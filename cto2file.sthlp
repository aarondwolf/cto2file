{smcl}
{* *! version 2.0.2 Aaron Wolf 9may2021}{...}
{title:Title}

{phang}
{cmd:cto2file} {hline 2} Convert a SurveyCTO XLSForm to an editable Microsoft Word document.

{marker syntax}{...}
{title:Syntax}

{pmore}
{cmd: cto2file} {cmd:using} {it:{help filename}} {cmd:,} {opt sa:ve(filename)} [{it:options}]

{* Using -help odkmeta- as a template.}{...}
{* 24 is the position of the last character in the first column + 3.}{...}
{synoptset 24 tabbed}{...}
{synopthdr}
{synoptline}

{syntab:Main}
{synopt:{opt sa:ve(filename)}*}name of the new .docx or .xlsx file to be saved. The extension determines the file type.{p_end}
{synopt:{opt def:ault(language)}}set default language for {it:label} columns with no language specified.{p_end}
{synopt:{opt l:anguage(language)}}set language to use for {it:label} and {it:hint} values.{p_end}
{synopt:{opt co:mmand(string)}}execute a set of commands after importing the {it:survey} sheet.{p_end}
{synopt:{opt t:itle(string)}}survey title for Word or Excel document.{p_end}
{synopt:{opt sub:title(string)}}survey subtitle for Word or Excel document.{p_end}
{synopt:{opt p:agesize(string)}} {help putdocx} page size for Word.{p_end}
{synopt:{opt int:length(int)}}number of spaces provided for {it:integer} variable inputs (Word only).{p_end}
{synopt:{opt dec:length(int)}}number of spaces provided for {it:decimal} variable inputs (Word only).{p_end}
{synopt:{opt max:heading(int)}}maximum number of nested group-levels to keep.{p_end}
{synopt:{opt keepr:elevance}}edits and displays {it:relevance} expressions in the Word document.{p_end}


{syntab:Variable Numbering}
{synopt:{opt novarnum}}do not re-number variables.{p_end}
{synopt:{opt noheadnum}}do not add auto-generated heading levels to begin and end group labels.{p_end}
{synopt:{opt omitn:ames(namelist)}}do not generate variable numbers for specified variable names.{p_end}
{synopt:{opt omitt:ypes(namelist)}}do not generate variable numbers for specified types.{p_end}
{synopt:{opt keept:ypes(namelist)}}apply new variable numbers to specified variable types.{p_end}


{syntab:Drop Variables}
{synopt:{opt dropg:roups(namelist)}}drop all group headings with the names specified.{p_end}
{synopt:{opt dropn:ames(namelist)}}drop all variables with the names specified.{p_end}
{synopt:{opt dropt:ypes(namelist)}}drop all variables with the types specified.{p_end}


{syntab:{it:Other} Variables}
{synopt:{opt others:uffix(string)}}specify a suffix common to all variables that represent "Other" text inputs.{p_end}
{synopt:{opt othern:ames(namelist)}}specify that all variables with the names specified are "Other" variables.{p_end}
{synopt:{opt keepo:ther}}do not drop "Other" variables in the document.{p_end}

{synoptline}
{p 4 6 2}* {opt sa:ve(filename)} is required.{p_end}


{title:Description}

{pstd}
{cmd:cto2file} uses the metadata generated by an existing SurveyCTO XLSForm
to construct an editable Microsoft Word document or Excel file with 
appropriately-formated tables containing the variable names or numbers, the 
question asked ({it:labels}), hints, skip patterns (via {it:relevance} 
expressions), and choice lists.

{pstd}
The resulting Word document contains headings for each group (with a cascading 
style using Word's default Heading 1, Heading 2,etc. for nested groups). The
Excel document will have custom styled headings for each group level. 

{p2colset 8 12 12 4}
{p2col:1.}The first column contains the variable {it:name}, or a calculated
variable number.{p_end}
{p2col:2.}The second column contains the information in {it:label} (or
{it:label}:{it:language}, if {opt l:anguage} is specified), hints, and skip 
patterns ({it:relevance} exressions).{p_end}
{p2col:3.}The third column contains a numbered list of possible choices 
(for {it:select_one} and {it:select_multiple} variables), a series of boxes to 
input digits for {it:integer} or {it:decimal} variables, 
the phrase "calculated value" for {it:calculate} or {it: calculate_here}
variables, and is blank otherwise. For {it:note} fields, the second and third
column are merged into one.{p_end}
{p2colreset}

{title:Remarks}

{pstd}
The command imports the {it:survey} sheets from the {help filename}
specified. It then runs a series of commands to prepare the metadata for
writing via the {help putdocx} command. These include: {p_end}
{p2colset 8 12 12 4}
{p2col:1.}Dropping all groups, names, and types specified in the {opt dropg:groups}, {opt dropn:ames}, and {opt dropt:ypes} options.{p_end}
{p2col:2.}Executing commands specified in {opt co:mmand}.{p_end}
{p2col:3.}Generates new question numbers prefixed by V (if {opt novarnum} is not specified). E.g. V001, V002, etc. {p_end}
{p2col 8 16 16 4:}{bf:Note}: {it:note} variables are not considered questions, and receive their own numbers prefixed by "__n".{p_end}
{p2col:4.}Generates heading levels (e.g. Heading 1.0, Heading 1.1, Heading 2.0, etc.) for nested groups.{p_end}
{p2col:5.}Re-writes {it:relevance} expressions to be more readable. This includes:{p_end}
{p2col 16 20 20 4:a.)}Converting {it:${name}} expressions to {it:[name]}.{p_end}
{p2col 16 20 20 4:b.)}Converting {it:selected(${name},'value')} to {it:[name]=value}.{p_end}
{p2col 16 20 20 4:c.)}Adding a variable number to the variable reference (if {opt novarnum} is not specified). E.g. V001 [name]. {p_end}
{p2col:6.}Uses {it:constraint} expressions to estimate the number of integers/decimal places in {it:integer} and {it:decimal} type variables. This defaults to {opt int:length} and {opt dec:length} for very complicated expressions.{p_end}
{p2col:7.}Create individual tables for each set in the {it:choices} sheet.{p_end}
{p2col:8.}Writes a well-formated table in Word using the {help putdocx} command for each row in the {it:survey} sheet.{p_end}
{p2col:9.}Saves the document to the {help filename} specified in {opt sa:ve()}.{p_end}
{p2colreset}

{pstd}
{cmd: cto2file} requires the {help putdocx} command introduced in Stata 15 to work.

{title:Options}

{dlgtab:Main}

{phang}{opt sa:ve(filename)} specifies the name of the new Word document to be saved. This must be a valid {help filename} with a .docx extension. This option is required.{p_end}

{phang}{opt def:ault(language)} sets the default language if there is a {it:label} or {it:hint} column with no language specified. The default is to rename this column {it:labelenglish}/{it:hintenglish}. {p_end}

{phang}{opt l:anguage(language)} set language to use for {it:label} and {it:hint} values. The default is the value from {opt def:ault()} option, which defaults to {it:english} if unspecified.{p_end}

{phang}{opt co:mmand(string)} executes a set of commands after importing the 
{it:survey} sheet. This option can allow you to, e.g., replace the labels select 
variables, or replace constraint expressions to simpler versions. Each command 
should be separated by compound double quotes and a space. E.g.:{p_end}

		{cmd: cto2file using example.xslx, save(newfile.docx)}	///
			{cmd: command(`"command1"' `"command2"' ... )}

{phang}{opt t:itle(string)} sets a survey title which appears at the beginning of the document.{p_end}

{phang}{opt sub:title(string)}sets a survey subtitle which apears beneath the title.{p_end}

{phang}{opt p:agesize(string)} sets the page size for the Word document. The default is {it:letter}. Any valid {help putdocx} page size is accepted.{p_end}

{phang}{opt scheme(string)} sets the color scheme used in the final document. 
This option is not related to Stata's {help schemes}. Available color schemes 
are {it:purple}, {it:blue}, {it:green}, {it:orange}, {it:drab}, and 
{it:colorful}. The default is {it:blue}.{p_end}

{phang}{opt int:length(int)}number of spaces provided for {it:integer} variable inputs when this cannot be determined by the variable constraints. Default is 9.{p_end}

{phang}{opt dec:length(int)}number of spaces provided for {it:decimal} variable inputs when this cannot be determined by the variable constraints. Default is 9.{p_end}

{phang}{opt split}create a new table in the Word document for every top-level group.{p_end}

{phang}{opt max:heading(int)}maximum number of nested group-levels to keep. Default is 3.{p_end}

{phang}{opt keepr:elevance}apply set of edits to {it:relevance} expressions and display them in the document.{p_end}


{dlgtab:Variable Numbering}

{phang}{opt novarnum}do not re-number variables. Keep the {it:name} value as the variable names in the left-most column of the document.{p_end}

{phang}{opt omitn:ames(namelist)}do not generate variable numbers for specified variable names.{p_end}

{phang}{opt omitt:ypes(namelist)}do not generate variable numbers for specified types. By default, {it:begin group}, {it:end group}, {it:begin repeat}, {it:end repeat}, {it:comments}, {it:note}, {it:username}, {it:caseid}, {it:start}, {it:end}, {it:deviceid}, {it:text audit}, and {it:audio audit} are omitted from having variable numbers generated. {p_end}

{phang}{opt keept:ypes(namelist)}apply new variable numbers to specified variable types.{p_end}


{dlgtab:Drop Variables}

{phang}{opt dropg:roups(namelist)}drop all group headings with the names specified.{p_end}

{phang}{opt dropn:ames(namelist)}drop all variables with the names specified.{p_end}

{phang}{opt dropt:ypes(namelist)}drop all variables with the types specified.{p_end}


{dlgtab:Other Variables}

{phang}{opt others:uffix(string)}specify a suffix common to all variables that represent "Other" or "Other (Specify)" text inputs. These will not have unique variable numbers added.{p_end}

{phang}{opt othern:ames(namelist)}specify that all variables with the names specified are "Other" or "Other (Specify)" variables.{p_end}

{phang}{opt keepo:ther}keep "Other" and "Other (Specify)" variables in the document. The default is to drop them.{p_end}





{title:Author}

{pstd}Aaron Wolf, Yale University {p_end}
{pstd}aaron.wolf@yale.edu{p_end}













