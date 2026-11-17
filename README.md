# Transaction Data Cleaning with Excel Power Query

## Brief One Line Summary

Transformed 14,038 messy, single-column transaction text records into a clean, analysis-ready table of four typed columns using Excel Power Query.

---

## Overview

This project demonstrates an end-to-end data cleaning workflow in Microsoft Excel, using Power Query as the primary tool. The raw data arrives as one long text string per row, combining a transaction number, client name, business status, a label, and a currency amount with a credit/debit marker. The goal was to parse each string into separate, correctly typed fields that can be filtered, summarised, and visualised in Excel, Power BI, or Tableau.

The workbook contains three sheets that show the full journey from raw to final output:

| Sheet | Purpose |
|---|---|
| `RAW DATA` | Original data: 14,038 rows in a single column (`Column1`) |
| `Table1` | Intermediate output after the first Power Query: four columns, amount still stored as text |
| `CLEANED DATASET` | Final output after the second Power Query: amount converted to a numeric value |

---

## Problem Statement

Business data exported from accounting or banking systems often arrives as unstructured text. In this dataset, every record was packed into one cell, which makes it impossible to sort, filter, aggregate, or chart without cleaning it first.

The specific problems in the raw data were:

1. **Everything in one column.** Transaction number, client, status, label, and amount were all merged into a single text string per row.
2. **Inconsistent leading whitespace.** Every one of the 14,038 rows begins with a varying number of leading spaces.
3. **Irregular internal spacing.** Extra spaces appear between the fields, and the number of spaces differs from row to row, so splitting on a fixed position is unreliable.
4. **Mixed delimiters.** The fields are separated by different markers: a space after the transaction number, a hyphen (`-`) between client and status, and a rupee symbol (`₹`) before the amount.
5. **Redundant label.** The word `Internal` appears in every row and carries no analytical value.
6. **Currency symbol and Indian number formatting.** Amounts use the rupee symbol and the lakh/crore comma grouping (for example `₹18,75,600.00`), so they are read as text rather than numbers.
7. **Credit/debit suffix attached to the amount.** Each amount ends with `CR` or `DR`, which prevents numeric conversion until it is separated.
8. **Wrong data types.** After splitting, the transaction number and amount need to be converted from text to whole number and decimal number respectively.

---

## Dataset

**Source file:** `data_cleaning_practice_1.xlsx`
**Size:** 14,038 records

### Raw data structure

The raw data sits in a single column named `Column1`. Each row follows this pattern:

```
<spaces> TRAN# <spaces> CLIENT NAME - STATUS <spaces> Internal ₹AMOUNT CR/DR
```

Examples of raw rows:

```
   500112345 AARAV INDUSTRIES - PRIVATE LIMITED  Internal ₹2,45,890.50 CR
        500445678 KRISHNA ALLOYS - PROPRIETOR  Internal ₹6,540.75 DR
 500334567 SHIVNERI INFRA DEVELOPERS - PARTNERSHIP  Internal ₹18,75,600.00 CR
```

### Fields extracted

| Column | Description | Final type |
|---|---|---|
| `TRAN#` | Transaction number | Whole number |
| `CIENTS` | Client (business) name | Text |
| `STATUS` | Type of business entity | Text |
| `AMOUNT.1` | Transaction amount in rupees | Decimal number |

> Note: the column header `CIENTS` is spelled as it appears in the workbook (it is a typo for `CLIENTS`).

### Dataset profile

| Measure | Value |
|---|---|
| Total rows | 14,038 |
| Distinct transaction records | 360 |
| Distinct client names | 358 |
| Business status categories | 4 |
| Credit (`CR`) transactions | 11,220 |
| Debit (`DR`) transactions | 2,818 |
| Transactions with zero amount | 1,735 |
| Amount range | 0 to 1,975,600 |

**Business status distribution**

| Status | Records |
|---|---|
| Private Limited | 4,762 |
| Partnership | 3,683 |
| Proprietor | 2,798 |
| LLP | 2,795 |

---

## Tools and Technologies

- **Microsoft Excel** (Microsoft 365 / Excel 2016 or later)
- **Power Query** (Get & Transform Data) for all transformations
- **Power Query M language** (steps generated through the editor interface)

---

## Methods

All cleaning was performed with Power Query, in two connected queries. Every step below is recorded in the Applied Steps pane, which makes the process transparent and repeatable.

### Query 1: `Table1` (parse the raw text into four columns)

| Step | Power Query action | Purpose |
|---|---|---|
| 1 | Load `RAW DATA` as a table and set `Column1` to Text | Establish a consistent starting data type |
| 2 | **Trim** | Remove the leading and trailing spaces from every row |
| 3 | **Split Column by Delimiter** (space, leftmost occurrence) | Separate the transaction number from the rest of the string |
| 4 | Change type to Whole Number | Convert the transaction number from text |
| 5 | **Split Column by Delimiter** (hyphen) | Separate the client name from the status and amount section |
| 6 | **Split Column by Delimiter** (`₹`) | Separate the status and label from the amount |
| 7 | **Trim** and **Clean** | Remove leftover spaces and non-printing characters |
| 8 | **Split Column by Delimiter** (space, rightmost occurrence) | Separate the status from the redundant `Internal` label |
| 9 | **Remove Columns** | Drop the `Internal` label column |
| 10 | **Rename Columns** | Assign meaningful names: `TRAN#`, `CIENTS`, `STATUS`, `AMOUNT` |

Result: four clean columns, with `AMOUNT` still stored as text such as `2,45,890.50 CR`.

### Query 2: `Table1_1` (convert the amount to a number)

| Step | Power Query action | Purpose |
|---|---|---|
| 1 | Set column types | Keep `TRAN#` as Whole Number and the other columns as Text |
| 2 | **Split Column by Delimiter** (space, rightmost occurrence) on `AMOUNT` | Separate the numeric portion from the `CR` / `DR` suffix |
| 3 | Change type to **Decimal Number** | Convert the amount from text to a numeric value, handling the Indian comma grouping |
| 4 | **Remove Columns** | Drop the `CR` / `DR` suffix column |

Result: the final `CLEANED DATASET` sheet, with `AMOUNT.1` stored as a true number.

---

## Key Insights

**What the cleaning revealed about the data**

- The raw export is uniformly structured, which is why delimiter-based splitting worked reliably across all 14,038 rows with no parsing failures.
- Private Limited companies account for the largest share of records (about 34%), followed by Partnerships (about 26%). Proprietors and LLPs are each about 20%.
- Credit transactions outnumber debit transactions by roughly four to one (11,220 versus 2,818).
- About 12% of records (1,735) carry a zero amount, which is worth reviewing with the data owner before any financial analysis.

**Data quality observations and next steps**

These are points I identified on review and would address in a further iteration:

- **Repeated records.** Although there are 14,038 rows, only 360 are distinct. The remaining rows are exact repeats, and they have been retained in the cleaned dataset. A `Remove Duplicates` step in Power Query would reduce the table to its distinct records if the repeats are not genuine separate transactions.
- **Credit/debit direction is not preserved.** The `CR` / `DR` suffix was removed to obtain a numeric amount, so the cleaned amount column holds absolute values only. Keeping the suffix as its own column, or converting debits to negative values, would retain this information for analysis.
- **Trailing spaces in text fields.** The `CIENTS` and `STATUS` values in the final output still end with a space. Applying a Trim to those two columns would make grouping and lookups more reliable.

---

## Dashboard/Model/Output

### Before and after

| Raw (`RAW DATA`) | Cleaned (`CLEANED DATASET`) |
|---|---|
| `   500112345 AARAV INDUSTRIES - PRIVATE LIMITED  Internal ₹2,45,890.50 CR` | 500112345 / AARAV INDUSTRIES / PRIVATE LIMITED / 245890.50 |
| `        500445678 KRISHNA ALLOYS - PROPRIETOR  Internal ₹6,540.75 DR` | 500445678 / KRISHNA ALLOYS / PROPRIETOR / 6540.75 |
| ` 500334567 SHIVNERI INFRA DEVELOPERS - PARTNERSHIP  Internal ₹18,75,600.00 CR` | 500334567 / SHIVNERI INFRA DEVELOPERS / PARTNERSHIP / 1875600.00 |

### Output structure

| TRAN# | CIENTS | STATUS | AMOUNT.1 |
|---|---|---|---|
| 500112345 | AARAV INDUSTRIES | PRIVATE LIMITED | 245,890.50 |
| 500223456 | MEHTA AGRO FOODS | PROPRIETOR | 0.00 |
| 500334567 | SHIVNERI INFRA DEVELOPERS | PARTNERSHIP | 1,875,600.00 |
| 500445678 | KRISHNA ALLOYS | PROPRIETOR | 6,540.75 |
| 500556789 | VARDHAN EXPORT HOUSE | PRIVATE LIMITED | 987,000.00 |

The final table is ready to be loaded into a PivotTable, Power BI, or Tableau for further analysis and visualisation.

---
## Results & Conclusion

- **Outcome:** 14,038 unstructured text rows were converted into a structured four-column table with correct data types (whole number, text, and decimal number), with no rows lost or rejected during parsing.
- **Approach:** the work was completed entirely with Power Query's built-in transformations (trim, clean, split by delimiter, change type, remove and rename columns), with no formulas or manual editing. Every step is recorded and can be re-run when new data arrives.
- **Skills demonstrated:** text parsing, delimiter-based splitting, data type conversion, handling regional number formats, documenting data quality issues, and building a repeatable cleaning process.
- **Next steps:** remove repeated records, retain the credit/debit indicator, trim the remaining spaces, and connect the cleaned table to a Power BI or Tableau dashboard.

---

## Connect with Me

**SAMANPREET KAUR**
BA Student | Aspiring Data Analyst, Business Analyst, and BI Developer

- LinkedIn: [www.linkedin.com/in/kaursamanpreet](www.linkedin.com/in/kaursamanpreet)
- Email: [mailto:kaursamanpreet970@outlook.com](mailto:kaursamanpreet970@outlook.com)
- GitHub: [https://github.com/kaursamanpreet970-pixe](https://github.com/kaursamanpreet970-pixel)
