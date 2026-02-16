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
- Logarithmic transformation of indemnity payments  

The goal is to ensure reproducible and transparent data processing.

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
  - `SeverityCode` (1–9 ordinal scale)  
  - `PostReform` indicator (0 = pre-reform)

---

### 2. ConstructMPLAnalysisDataset_Post (Post-Reform)

Processes claims resolved on or after the reform cutoff date (March 24, 2023).

Functionality:

- Performs identical extraction and cleaning steps as the pre-reform macro  
- Automatically assigns `PostReform = 1`

Both macros assume consistent FLOIR export structure.

---

### 3. Logarithmic Transformation of Indemnity Payments

After dataset construction, indemnity payments are log-transformed within the `Clean` sheet to prepare the data for statistical analysis.

#### Rationale

Medical malpractice indemnity payments exhibit substantial positive skew due to the presence of extreme high-value settlements. To reduce the influence of these outliers and stabilize variance across injury severity groups, a base-10 logarithmic transformation is applied.

#### Process

- Reads Indemnity Paid from Column D of the `Clean` sheet  
- Computes `log10(indemnity)` for each positive value  
- Writes the result to Column G as `Log_Indemnity`  
- Preserves original indemnity values  

#### Formula Applied


#### Example Conversions

| Indemnity ($) | Log10 Value |
|---------------|------------|
| 250,000       | 5.398      |
| 1,000,000     | 6.000      |
| 10,000,000    | 7.000      |

Each increase of 1 unit on the log scale corresponds to a tenfold increase in indemnity payment size.

This transformation improves compatibility with ANOVA assumptions by reducing skewness and compressing extreme upper-tail payouts.

---

## Required Excel Structure

The workbook must contain:

### Sheet Names

- `Raw` (contains FLOIR export)  
- `Clean` (analysis output sheet)

### Required Column Positions in `Raw`

| Variable | Column |
|----------|--------|
| Claim Number | B |
| Injury Severity | O |
| Indemnity Paid | X |
| Final Disposition Date | AA |

If column positions differ, the macro must be adjusted accordingly.

---

## Usage

1. Paste the FLOIR export into the `Raw` sheet.  
2. Run the appropriate preprocessing macro.  
3. Apply the log transformation macro to the `Clean` sheet.  
4. Export or analyze the processed dataset as needed.

---

## Reproducibility and Transparency

These scripts were used to prepare datasets for empirical analysis of indemnity payments across injury severity levels and reform periods.

All transformations applied to the data are documented in this repository.

This repository contains code only. No raw claim data are included.


