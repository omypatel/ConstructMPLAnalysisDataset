# Florida MPL Closed Claim Data Processing

This repository contains the VBA macro used to construct the analysis-ready dataset from Florida Office of Insurance Regulation (FLOIR) Medical Professional Liability closed-claim exports.

## Functionality
The macro:
- Extracts claim number, injury severity, indemnity paid, and final disposition date
- Standardizes monetary values
- Codes injury severity into ordinal categories (1–9)
- Generates a binary post-reform indicator (cutoff: March 24, 2023)

## Required Excel Structure
- Sheet names: "Raw" and "Clean"
- Claim Number column: B
- Injury Severity column: O
- Indemnity column: X
- Final Disposition column: AA

## Usage
1. Paste FLOIR export into "Raw"
2. Run ConstructMPLAnalysisDataset
3. Analysis-ready dataset appears in "Clean"
