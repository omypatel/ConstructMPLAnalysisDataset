# Florida MPL Closed Claim Data Processing

This repository contains VBA macros used to construct analysis-ready datasets from the Florida Office of Insurance Regulation (FLOIR) Medical Professional Liability (MPL) Closed Claim Database.

The scripts standardize raw export files and generate structured datasets suitable for statistical analysis of indemnity payments across injury severity levels and reform periods.

---

## Purpose

FLOIR closed-claim exports include numerous administrative fields and formatting inconsistencies that are not analysis-ready. These VBA macros automate:

- Extraction of key analytical variables
- Standardization of monetary values
- Normalization of date formats
- Ordinal coding of injury severity (1–9 scale)
- Generation of reform-period indicators

The goal is to ensure reproducible, transparent data processing.

---

## Macros Included

### 1. ConstructMPLAnalysisDataset (Pre-Reform)

Processes claims resolved prior to the 2023 tort reform cutoff date.

Functionality:
- Extracts Claim Number (Column B)
- Extracts Injury Severity (Column O)
- Extracts Indemnity Paid (Column X)
- Extracts Final Disposition Date (Column AA)
- Cleans monetary formatting
- Generates:
  - SeverityCode (1–9 ordinal scale)
  - PostReform indicator (0 = pre-reform)

---

### 2. ConstructMPLAnalysisDataset_Post (Post-Reform)

Processes claims resolved on or after the reform cutoff date (March 24, 2023).

Functionality:
- Performs identical extraction and cleaning steps as the pre-reform macro
- Automatically assigns PostReform indicator = 1

Both macros assume consistent FLOIR export structure.

---

## Required Excel Structure

The workbook must contain:

Sheet Names:
- "Raw" (contains FLOIR export)
- "Clean" (analysis output sheet)

Required Column Positions in "Raw":
- Claim Number → Column B
- Injury Severity → Column O
- Indemnity Paid → Column X
- Final Disposition Date → Column AA

If column positions differ, the macro must be adjusted accordingly.

---

## Usage

1. Paste the FLOIR export into the "Raw" sheet.
2. Run the appropriate macro.
3. The cleaned dataset will populate the "Clean" sheet.
4. Export or analyze the cleaned dataset as needed.

---

## Reproducibility and Transparency

These scripts were used to prepare datasets for empirical analysis of indemnity payments across injury severity levels and reform periods.

This repository contains code only. No raw claim data are included.

---

## Version

Version 1.0  
Reform Cutoff Date: March 24, 2023  
Author: Omy Patel
