# cto2file

---

*cto2file* is a Stata user-written command that converts a SurveyCTO-compatible XLSForm into a readable and editable Word document (.docx) or Excel workbook (.xlsx).

This command is very much an early work in progress. To date, I have only used it myself on a few SurveyCTO forms for my own projects. It works for these, but consistently throws errors with every new form. You are welcome to use this command yourself (or, even better, start a branch and try to improve it!), but be warned that it will be very finicky.

The command utilizes Stata's *putdocx* command, and thus requires Version 15 or above to operate.

## Installing via *net install*

The current version is still a work in progress. To install, user can use the net install command to download from the project's Github page:

```
net install cto2file, from("https://aarondwolf.github.io/cto2file")
```

## Syntax

```
    cto2file using filename , save(filename) [options]
```

## Description

cto2file uses the metadata generated by an existing SurveyCTO XLSForm to construct an editable Microsoft Word document or Excel file with appropriately-formated tables containing the variable names or numbers, the question asked (labels), hints, skip patterns (via relevance expressions), and choice lists.

The resulting Word document contains headings for each group (with a cascading style using Word's default Heading 1, Heading 2,etc. for nested groups). The Excel document will have custom styled headings for each group level.

   1.  The first column contains the variable name, or a calculated variable number.
      2.  The second column contains the information in label (or label:language, if language is specified), hints, and skip patterns (relevance exressions).
      3.  The third column contains a numbered list of possible choices (for select_one and select_multiple variables), a series of boxes to input digits for integer or decimal variables, the phrase "calculated value" for calculate or calculate_here variables, and is blank otherwise. For note fields, the second and third column are merged into one.

## Remarks

The command imports the survey sheets from the filename specified. It then runs a series of commands to prepare the metadata for writing via the putdocx command. These include:

      1.  Dropping all groups, names, and types specified in the dropggroups, dropnames, and droptypes options.
   2.  Executing commands specified in command.
      3.  Generates new question numbers prefixed by V (if novarnum is not specified). E.g. V001, V002, etc. Note: note variables are not considered questions, and receive their own numbers prefixed by "__n".
      4.  Generates heading levels (e.g. Heading 1.0, Heading 1.1, Heading 2.0, etc.) for nested groups.
   5.  Re-writes relevance expressions to be more readable. This includes:
           a.) Converting ${name} expressions to [name].
           b.) Converting selected(${name},'value') to [name]=value.
           c.) Adding a variable number to the variable reference (if novarnum is not specified). E.g. V001 [name].
      6.  Uses constraint expressions to estimate the number of integers/decimal places in integer and decimal type variables. This defaults to intlength and declength for very complicated expressions.
   7.  Create individual tables for each set in the choices sheet.
      8.  Writes a well-formated table in Word using the putdocx command for each row in the survey sheet. If a .xslx extension is specified, writes a well-formated Excel table into the specified worksheet.
   9.  Saves the document to the filename specified in save().

cto2file requires the putdocx command introduced in Stata 15 to work when using .docx as a file extension.

## Options

### Main

**save(filename)** specifies the name of the new Word document to be saved. This must be a valid filename with a .docx extension. This option is required.

**default(language)** sets the default language if there is a label or hint column with no language specified. The default is to rename this column labelenglish/hintenglish.

**language(language)** set language to use for label and hint values. The default is the value from default() option, which defaults to english if unspecified.

**command(string)** executes a set of commands after importing the survey sheet. This option can allow you to, e.g., replace the labels select variables, or replace constraint expressions to simpler versions. Each command should be separated by compound double quotes and a space. E.g.:

```
             cto2file using example.xslx, save(newfile.docx)        ///
                     command(`"command1"' `"command2"' ... )
```

**title(string)** sets a survey title which appears at the beginning of the document.

**subtitle(string)** sets a survey subtitle which appears beneath the title.

**pagesize(string)** sets the page size for the Word document. The default is letter. Any valid putdocx page size is accepted.

**scheme(string)** sets the color scheme used in the final document.  This option is not related to Stata's schemes. Available color schemes are purple, blue, green, orange, drab, and  colorful. The default is blue.

**intlength(int)** number of spaces provided for integer variable inputs when this cannot be determined by the variable constraints. Default is 9.

**declength(int)** number of spaces provided for decimal variable inputs when this cannot be determined by the variable constraints. Default is 9.

**splitcreate** a new table in the Word document for every top-level group.

**maxheading(int)** maximum number of nested group-levels to keep. Default is 3.

**keeprelevance** apply set of edits to relevance expressions and display them in the document.

### Variable Numbering

**novarnum** do not re-number variables. Keep the name value as the variable names in the left-most column of the document.

**omitnames(namelist)** do not generate variable numbers for specified variable names.

**omittypes(namelist)** do not generate variable numbers for specified types. By default, begin group, end group, begin repeat, end repeat, comments, note, username, caseid, start, end, deviceid, text audit, and audio audit are omitted from having variable numbers generated.

**keeptypes(namelist)** apply new variable numbers to specified variable types.

### Drop Variables

**dropgroups(namelist)** drop all group headings with the names specified.

**dropnames(namelist)** drop all variables with the names specified.

**droptypes(namelist)** drop all variables with the types specified.

### Other Variables

**othersuffix(string)** specify a suffix common to all variables that represent "Other" or "Other (Specify)" text inputs. These will not have unique variable numbers added.

**othernames(namelist)** specify that all variables with the names specified are "Other" or "Other (Specify)" variables.

**keepotherkeep** "Other" and "Other (Specify)" variables in the document. The default is to drop them.

## Author

Aaron Wolf, Northwestern University
aaron.wolf@u.northwestern.edu