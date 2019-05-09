# cto2file
*cto2file* is a Stata user-written command that converts a SurveyCTO-compatible XLSForm into a readable and editable Word document (.xlsx).

This command is very much an early work in progress. To date, I have only used it myself on a few SurveyCTO forms for my own projects. It works for these, but consistently throws errors with every new form. You are welcome to use this command yourself (or, even better, start a branch and try to improve it!), but be warned that it will be very finicky.

The command utilizes Stata's *putdocx* command, and thus requires Version 15 or above to operate.


## Current Options
The command accepts a number of options:
- *default*: Specifies the default language for the form. This is used when the form has a default language that is not included in the *label*, *hint*, or *constraint message* fields. For example, if your default language for the form is english, and you also have the form translated into hindi, you may have a *label* field as well as a *label:hindi* field. In this case, you do **not** need to specify a default language. However, if you do not have a blank *label* field, but instead have both a *label:english* and *labelchichewa* field, you should specify which one is the default language. If no language is specified as the default, english will be assumed.
- *language*: This specifies which language you want to use to construct the Word document. As in the example above, if no *default* language or *language* options are specified, the command will assume english, and use the un-specified language field.
- *dropgroups*: Drops the *begin group* and *end group* fields with the names specified. Does not drop fields within the groups themselves, only the group identifiers.
- *dropames*: Drops specific names from the Word document.
- *droptypes*: Drops all names with a specific type. Any valid SurveyCTO type is accepted (e.g. *droptypes(calculate_here)*)
- *omitnames*: Keeps the name in the final Word Document, but does not assign it a question number (or display the SurveyCTO name).
- *omittypes*: Keeps all names with these types in the final Word Document, but does not assign a variable number or display teh SurveyCTO name. By default, groups, setup variables, and notes all have omitted variable numbers/names.
- *keeptypes*: Allows you to specify types to display the name (e.g. caseid). Both omittypes and keeptypes cannot be included simultaneously.
- *othersuffix*: If you have a consistent suffix for all "Other" variables (text variables which trigger if "Other" is selected), you can specify that suffix here, and it will format those fields appropriately, using the variable number from the assumed original variable.
- *othernames*: Allows you to specify individual fields that represent "Other" options for a preceding select_one or select_multiple field.
- *keepther*: By default, all "Other" variables are dropped for concision. Specifying this option will add them.
- *keeprelevance*: By default, relevance fields are not displayed. Specifying this option displays them underneath the variable name in italics.
- *title*: Specifies a title for the document (E.g. ABC Survey)
- *subtitle*: Specifies a sub-title for the document (E.g. "v. 2019.01.01")
- *maxheading*: The command generates heading levels based on nested groups. So the top-level group is heading 1, the first nested group is heading 2, etc. So the first-top level group will have heading number 1. The second top-level group would have heading 2. The first nested group within heading 2 will have the heading 2.1. This option allows the user to specify the maximum number of heading levels to be added to the document. The default is 5 (e.g. a group could be 1.1.1.1.1). After this point, the group name will be omitted from the survey, but the fields within them will remain.
- *intlength*: Integer fields display boxes that can be written in, one for each unit. The command uses constraints to estimate how many digits are needed. If a field's constraint is something like ".>=0 or .=-999", the command will recognize the maximum possible number of digits as 4: 1 for the "-", and three for the "999". This integer will display with 4 boxes. The *intlength* option allows the user to set a default number for fields with no constraints, or complex constraints which may cause the command to fail to accurately parse the number of digits. The default is 9.
- *declength*: Similar to intlength, but adding a decimal and digits after the decimal.
- *scheme*: [NOT YET IMPLEMENTED] Allows the user to apply different sets of colors, fonts, and text sizes from a selection of pre-set schemes. This is not related to Stata's *scheme* command, which sets schemes for graphs.
- *pagesize*: Sets the page size. Any *putdocx* page size is valid. The default is *letter*.
- *split*: Creates a separate table for group (with heading level 1). By default, the document is written as one big table.
- *novarnum*: Does not label variables with new variable numbers. Instead, used the *names* as variable numbers.
- *command*: Allows the user to specify any commands to eliminate or change specific rows from the "survey" sheet in the XLSForm prior to writing.
