# surveyctocompare

A Stata command to compare two versions of a SurveyCTO XLSForm and generate a clean, formatted Excel dashboard summarizing all changes in variables, labels, and choice lists.

## Features

- Detects added, removed, and updated survey questions
- Classifies label changes by nature (significant, minor, formatting)
- Compares choice lists and maps them to affected variables
- Exports a polished Excel dashboard with summaries and details
- Optional formatting using Python's openpyxl

## Installation

To install the latest version using the `github` package by E. F. Haghish:

```stata
github install kellymontano/surveyctocompare
```

To update the command in the future:

```stata
github update kellymontano/surveyctocompare
```

Alternatively, you can use classic `net install`:

```stata
net install surveyctocompare, from("https://raw.githubusercontent.com/kellymontano/surveyctocompare/main/") replace
```

## Syntax

```stata
surveyctocompare , form1("form_v1.xlsx") form2("form_v2.xlsx") output("dashboard.xlsx") [noformat]
```

## Options

- `form1()` and `form2()` — Paths to the two `.xlsx` SurveyCTO forms to compare  
- `output()` — File path for the generated Excel report  
- `noformat` — Skips optional Excel formatting (no Python required)

## Output

The Excel dashboard includes:

- A `Summary` sheet with counts by type of change
- `Variable Updates`: new or removed survey fields
- `Label Updates`: textual label changes with classification
- `Choice Updates`: added/removed/edited choices and affected variables

## Requirements

- Stata 16 or newer
- Optional: Python 3 with `openpyxl` installed (for formatting)

## Author

Kelly Montaño  
Innovations for Poverty Action (IPA)

---

Feel free to share, reuse, or contribute.
