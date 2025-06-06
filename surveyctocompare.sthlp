{smcl}
{* *! version 1.0.0 Kelly Montaño 2025-06-06}{...}
{cmd:help surveyctocompare}

{title:Title}

{hi:surveyctocompare} {hline 2} Compare two versions of a SurveyCTO XLSForm and create a formatted Excel dashboard

{title:Syntax}

{cmd:surveyctocompare} {cmd:,}
{opt form1("filename.xlsx")}
{opt form2("filename.xlsx")}
{opt output("output.xlsx")}
[{opt noformat}]

{title:Description}

{pstd}
{cmd:surveyctocompare} compares two versions of a SurveyCTO XLSForm (.xlsx) and creates a structured Excel dashboard that summarizes changes in variables, labels, and choice lists. Optionally, it formats the output using Python's openpyxl.

{title:Options}

{phang}
{opt form1("filename.xlsx")} specifies the path to the first form version.

{phang}
{opt form2("filename.xlsx")} specifies the path to the second form version.

{phang}
{opt output("output.xlsx")} sets the path and filename of the output Excel report.

{phang}
{opt noformat} skips the optional Python formatting step.

{title:Requirements}

{pstd}
This command requires Stata 16 or newer. Optional formatting requires a Python installation with the {bf:openpyxl} library.

{title:Author}

{pstd}
Kelly Montaño{break}
Innovations for Poverty Action (IPA)

{title:Version}

{pstd}
Version 1.0.0 — June 6, 2025

{title:See Also}

{pstd}
Installation:
{cmd:github install kellymontano/surveyctocompare}
