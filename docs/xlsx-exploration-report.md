# XLSX Exploration Report

**Generated:** 2026-01-31
**File:** `DSCR_NoArrayFormulas_Testing.xlsx`
**Story:** 2.1 - Explore xlsx Structure and Define Named Range Strategy

---

## 1. Sheets Overview (12 total)

| # | Sheet Name | Dimensions | Purpose |
|---|------------|------------|---------|
| 0 | **Input** | A1:Z996 | Main input sheet - user inputs |
| 1 | **API_Output** | A1:AK1000 | Main output sheet with G5:AK32 table |
| 2 | Eligibility | A1:Z1025 | Eligibility calculation sheet |
| 3 | LTV_Matrix | B1:P1000 | LTV lookup matrix |
| 4 | Product_Caps | A1:Z1011 | Product constraints/caps |
| 5 | Adjustments | A1:Z1001 | Price adjustments |
| 6 | Rate_Sheet_Data | A1:Z1000 | Rate sheet data |
| 7 | Pricing_Output | A1:Z1000 | Pricing calculations |
| 8 | Rate_Sheet_Import | A1:S70 | Rate import data |
| 9 | Claude Log | A1:F96 | Development log |
| 10 | README | A1:C130 | Documentation |
| 11 | API_Documentation | A1:D93 | API documentation |

**Focus Sheets:** Input, API_Output

---

## 2. Named Ranges (10 total - Updated 2026-01-31)

### Summary by Type

| Type | Count | Examples |
|------|-------|----------|
| SINGLE | 5 | FicoScore, LoanAmount, LoanEligiblity |
| HORIZONTAL | 1 | LiabilitiesHorizontal (1x4) |
| VERTICAL | 1 | LiabilitiesVertical (4x1) |
| TABLE | 3 | RateStackTable, PriceAdjustmentTable, EligiblityFailureReason |

### Summary by Classification

| Classification | Count |
|----------------|-------|
| INPUT | 4 | (SINGLE: 2, HORIZONTAL: 1, VERTICAL: 1) |
| OUTPUT | 6 | (SINGLE: 3, TABLE: 3) |

**Note:** TABLE type is always OUTPUT per strategy decision.

### Detailed Named Range List

#### INPUT Named Ranges

| Name | Type | Reference | Dimensions |
|------|------|-----------|------------|
| `FicoScore` | SINGLE | Input!$D$6 | 1x1 |
| `LoanAmount` | SINGLE | Input!$D$7 | 1x1 |
| `LiabilitiesHorizontal` | HORIZONTAL | Input!$D$65:$G$65 | 1 row x 4 cols |
| `LiabilitiesVertical` | VERTICAL | Input!$C$60:$C$63 | 4 rows x 1 col |

#### OUTPUT Named Ranges

| Name | Type | Reference | Formula/Notes |
|------|------|-----------|---------------|
| `LiabilitiesHorizontalSum` | SINGLE | Input!$D$70 | `=SUM(D65:H65)` |
| `LiabilitiesVerticalSum` | SINGLE | Input!$D$69 | `=SUM(C60:C63)` |
| `LoanEligiblity` | SINGLE | API_Output!$B$4 | `=Eligibility!D67` |
| `RateStackTable` | TABLE | API_Output!$G$5:$AK$32 | 28×31, 837 formula cells (100%) |
| `PriceAdjustmentTable` | TABLE | API_Output!$A$7:$C$22 | 16×3, 28 formula cells (58%) |
| `EligiblityFailureReason` | TABLE | API_Output!$A$30:$B$92 | 63×2, 124 formula cells (98%) |

---

## 3. Input Sheet Analysis

### Table Areas Identified

#### A7:C23 Area
- **Content:** Labels and input field names
- **Formulas:** 0 (no formulas)
- **Sample data:**
  - Row 7: Loan Amount
  - Row 8: Property Value
  - Row 9: LTV (Calculated)
  - Row 10: Cash Out Amount
  - Row 11: DSCR (Optional)
  - Row 13: PROPERTY INFORMATION header
  - Row 14: Property Type
  - Row 15: Unit Count

#### A30:B92 Area
- **Content:** Labels for loan features section
- **Formulas:** 0 (no formulas)
- **Sample data:**
  - Row 35: LOAN FEATURES header

#### Test Input Arrays (C60:C63 and D65:G65)
- **LiabilitiesVertical (C60:C63):** 4 rows x 1 col - vertical input array
- **LiabilitiesHorizontal (D65:G65):** 1 row x 4 cols - horizontal input array
- **Purpose:** Test named ranges for array inputs

---

## 4. API_Output Sheet Analysis

### Main Output Table: G5:AK32

#### Dimensions
- **Rows:** 5 to 32 = 28 rows
- **Columns:** G to AK = 31 columns (G=7, AK=37)
- **Total cells:** 868

#### Cell Composition
| Cell Type | Count | Percentage |
|-----------|-------|------------|
| Formulas | 837 | 96.4% |
| Static values | 31 | 3.6% |
| Empty | 0 | 0% |

#### Header Row (Row 5) - Static Values
First 10 columns:
- `rate`, `basePrice`, `dscrPriceAdj`, `totalPriceAdj`, `maxPrice`
- `isPriceEligible`, `monthlyPI`, `dscr`, `reserveMonths`, `maxLTV`

Last 10 columns:
- `chk_RefiNoDSCR`, `chk_MtgLates`, `chk_LTV80pct`, `chk_40yrIO_LTV`
- `chk_40yrIO_FICO`, `chk_2to4UnitDSCR`, `chk_2to4UnitLTV`
- `chk_Condo_Purchase`, `chk_Condo_Refi`, `priceFailureReason`

#### Data Rows (6-32) - All Formulas
- All 27 data rows contain formulas
- Formulas reference `Pricing_Output` sheet
- Example: `G6: =Pricing_Output!B27`

---

## 5. Key Findings for Named Range Strategy

### What Already Exists
1. **Single INPUT:** `FicoScore`, `LoanAmount` - working examples
2. **Single OUTPUT:** `LoanEligiblity` - working example
3. **Horizontal INPUT:** `LiabilitiesHorizontal` (1x4) - working example
4. **Vertical INPUT:** `LiabilitiesVertical` (4x1) - working example
5. **Calculated OUTPUTs:** `LiabilitiesHorizontalSum`, `LiabilitiesVerticalSum` - test formulas

### What Needs to be Created
1. **TABLE OUTPUT:** Named range for G5:AK32 (or G6:AK32 data-only)
2. **Horizontal OUTPUT:** Test case for horizontal array output
3. **Vertical OUTPUT:** Test case for vertical array output

### Classification Logic Validation
The current approach works:
- **INPUT:** First cell of range has NO formula
- **OUTPUT:** First cell of range HAS formula

### Open Questions
1. Should G5:AK32 include headers (row 5) or just data (G6:AK32)?
2. Should horizontal/vertical arrays be flattened to 1D in JSON?
3. How to handle empty cells in ranges?
4. How to handle formula errors (#N/A, #REF!) in output?

---

## 6. Recommended Named Range Naming Convention

Based on existing patterns:

| Range Type | Classification | Naming Pattern | Example |
|------------|----------------|----------------|---------|
| Single | INPUT | `<FieldName>` | `LoanAmount` |
| Single | OUTPUT | `<FieldName>` | `LoanEligiblity` |
| Horizontal | INPUT | `<FieldName>Horizontal` | `LiabilitiesHorizontal` |
| Horizontal | OUTPUT | `<FieldName>Horizontal` | `RatesHorizontal` |
| Vertical | INPUT | `<FieldName>Vertical` | `LiabilitiesVertical` |
| Vertical | OUTPUT | `<FieldName>Vertical` | `YearsVertical` |
| Table | OUTPUT | `<FieldName>Table` | `PricingTable` |

**Note:** Classification (INPUT vs OUTPUT) is determined by formula presence, not naming.

---

## Appendix: Script Used

Location: `scripts/explore_xlsx.py`

```bash
python scripts/explore_xlsx.py
```
