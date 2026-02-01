# Named Range Strategy

**Document:** Named Range Strategy for Multi-Cell Outputs
**Story:** 2.1 - Explore xlsx Structure and Define Named Range Strategy
**Last Updated:** 2026-01-31

---

## Overview

This document defines the naming conventions and handling strategy for named ranges in the DSCR calculation engine, covering single cells, arrays, and tables.

---

## Range Types

| Type | Dimensions | Can Be INPUT? | Can Be OUTPUT? | JSON Format |
|------|------------|---------------|----------------|-------------|
| SINGLE | 1×1 | Yes | Yes | Scalar value |
| HORIZONTAL | 1×N | Yes | Yes | 1D array `[v1, v2, ...]` |
| VERTICAL | N×1 | Yes | Yes | 1D array `[v1, v2, ...]` |
| TABLE | N×M | **No** | Yes | Object with headers + 2D data |

---

## Classification Logic

### Detection by Dimensions

```javascript
function getRangeType(parsed) {
  const rows = parsed.endRow - parsed.startRow + 1;
  const cols = colToNum(parsed.endCol) - colToNum(parsed.startCol) + 1;

  if (rows === 1 && cols === 1) return 'SINGLE';
  if (rows === 1 && cols > 1) return 'HORIZONTAL';
  if (rows > 1 && cols === 1) return 'VERTICAL';
  if (rows > 1 && cols > 1) return 'TABLE';
}
```

### Classification as INPUT vs OUTPUT

```javascript
function classifyNamedRange(model, namedRange) {
  const parsed = parseCellReference(namedRange.formula);
  const rangeType = getRangeType(parsed);

  // TABLE is always OUTPUT (INPUT tables not supported)
  if (rangeType === 'TABLE') {
    return 'OUTPUT';
  }

  // For SINGLE, HORIZONTAL, VERTICAL: check first cell for formula
  const hasFormula = isFormulaCell(model, sheetIndex, parsed.startRow, parsed.startCol);
  return hasFormula ? 'OUTPUT' : 'INPUT';
}
```

| Range Type | Classification Rule |
|------------|---------------------|
| SINGLE | First cell has formula → OUTPUT, else INPUT |
| HORIZONTAL | First cell has formula → OUTPUT, else INPUT |
| VERTICAL | First cell has formula → OUTPUT, else INPUT |
| TABLE | **Always OUTPUT** (INPUT tables not supported) |

---

## JSON Format Specifications

### SINGLE (1×1)

**INPUT Example:**
```json
{
  "inputs": {
    "LoanAmount": 900000
  }
}
```

**OUTPUT Example:**
```json
{
  "outputs": {
    "LoanEligiblity": "YES"
  }
}
```

### HORIZONTAL (1×N) - Flattened to 1D Array

**INPUT Example:**
```json
{
  "inputs": {
    "MonthlyRates": [0.05, 0.06, 0.07, 0.08]
  }
}
```

**OUTPUT Example:**
```json
{
  "outputs": {
    "QuarterlyTotals": [25000, 28000, 31000, 34000]
  }
}
```

### VERTICAL (N×1) - Flattened to 1D Array

**INPUT Example:**
```json
{
  "inputs": {
    "Liabilities": [500, 300, 200, 100]
  }
}
```

**OUTPUT Example:**
```json
{
  "outputs": {
    "YearlyProjections": [100000, 105000, 110000, 115000]
  }
}
```

### TABLE (N×M) - Structured Object

**OUTPUT Example:**
```json
{
  "outputs": {
    "PricingTable": {
      "type": "TABLE",
      "rows": 27,
      "cols": 31,
      "headers": ["rate", "basePrice", "dscrPriceAdjustment", "..."],
      "data": [
        [0.065, 100.0, -0.5, "..."],
        [0.070, 99.5, -0.5, "..."],
        [0.075, 99.0, -0.5, "..."]
      ]
    }
  }
}
```

**Key Points for TABLE:**
- First row of range = headers (auto-extracted)
- Remaining rows = data
- Empty cells represented as `null`
- All columns guaranteed same length (defined by range bounds)

---

## Naming Conventions

### Recommended Patterns

| Range Type | Classification | Pattern | Example |
|------------|----------------|---------|---------|
| Single | INPUT | `<FieldName>` | `LoanAmount`, `FicoScore` |
| Single | OUTPUT | `<FieldName>` | `LoanEligiblity` |
| Horizontal | INPUT | `<FieldName>` | `MonthlyRates` |
| Horizontal | OUTPUT | `<FieldName>` | `QuarterlyTotals` |
| Vertical | INPUT | `<FieldName>` | `Liabilities` |
| Vertical | OUTPUT | `<FieldName>` | `YearlyProjections` |
| Table | OUTPUT | `<FieldName>Table` | `PricingTable` |

**Notes:**
- Classification (INPUT vs OUTPUT) determined by formula detection, not naming
- `Table` suffix recommended for TABLE types to make intent clear
- No special prefixes required - type detected from dimensions

---

## Existing Named Ranges (as of 2026-01-31)

### INPUT Ranges

| Name | Type | Reference | Dimensions |
|------|------|-----------|------------|
| `FicoScore` | SINGLE | Input!$D$6 | 1×1 |
| `LoanAmount` | SINGLE | Input!$D$7 | 1×1 |
| `LiabilitiesHorizontal` | HORIZONTAL | Input!$D$65:$G$65 | 1×4 |
| `LiabilitiesVertical` | VERTICAL | Input!$C$60:$C$63 | 4×1 |

### OUTPUT Ranges

| Name | Type | Reference | Dimensions |
|------|------|-----------|------------|
| `LoanEligiblity` | SINGLE | API_Output!$B$4 | 1×1 |
| `LiabilitiesHorizontalSum` | SINGLE | Input!$D$70 | 1×1 |
| `LiabilitiesVerticalSum` | SINGLE | Input!$D$69 | 1×1 |

### TABLE OUTPUT Ranges (Added 2026-01-31)

| Name | Type | Reference | Dimensions | Description |
|------|------|-----------|------------|-------------|
| `RateStackTable` | TABLE | API_Output!$G$5:$AK$32 | 28×31 | Main pricing/rate stack output (837 formula cells) |
| `PriceAdjustmentTable` | TABLE | API_Output!$A$7:$C$22 | 16×3 | Price adjustment factors |
| `EligiblityFailureReason` | TABLE | API_Output!$A$30:$B$92 | 63×2 | Eligibility failure reasons (124 formula cells) |

### HORIZONTAL/VERTICAL OUTPUT Ranges

No natural HORIZONTAL or VERTICAL OUTPUT arrays exist in this spreadsheet. The system supports them, but this model doesn't have any. All multi-cell outputs are TABLE format.

**Supported but not present:**
- HORIZONTAL OUTPUT (1×N with formulas) - Not found in current model
- VERTICAL OUTPUT (N×1 with formulas) - Not found in current model

---

## Implementation Notes

### Setting INPUT Arrays

For HORIZONTAL and VERTICAL INPUT arrays, the API accepts a 1D array and maps values to cells:

```javascript
// HORIZONTAL: Input!$D$65:$G$65 (1 row × 4 cols)
// Values map left-to-right: D65, E65, F65, G65
setInputArray(model, "MonthlyRates", [0.05, 0.06, 0.07, 0.08]);

// VERTICAL: Input!$C$60:$C$63 (4 rows × 1 col)
// Values map top-to-bottom: C60, C61, C62, C63
setInputArray(model, "Liabilities", [500, 300, 200, 100]);
```

**Validation:**
- Array length MUST match range size
- Partial arrays not supported (all cells must be filled)

### Reading OUTPUT Arrays

```javascript
// HORIZONTAL/VERTICAL: Return as 1D array
getOutputValue(model, "QuarterlyTotals"); // [25000, 28000, 31000, 34000]

// TABLE: Return as structured object
getOutputValue(model, "PricingTable");
// { headers: [...], data: [[...], [...], ...] }
```

### Empty Cell Handling

| Scenario | Representation |
|----------|----------------|
| Empty cell in array | `null` |
| Empty cell in table | `null` |
| Formula error (#N/A, #REF!) | `null` or error string (TBD) |

---

## IronCalc API Findings

### getDefinedNameList()

Returns array of named range objects:
```javascript
[
  { name: "FicoScore", formula: "Input!$D$6", scope: null },
  { name: "RateStackTable", formula: "API_Output!$G$5:$AK$32", scope: null },
  // ...
]
```

- `name`: Named range identifier
- `formula`: Cell reference with `$` notation (e.g., `Sheet!$A$1:$B$10`)
- `scope`: Sheet index (number) for sheet-scoped names, `null` for workbook-scoped

### parseCellReference() Compatibility

The existing `parseCellReference()` function correctly parses all IronCalc reference formats:

| Format | Example | Parsed |
|--------|---------|--------|
| Single cell | `Input!$D$6` | `{sheet:'Input', startCol:'D', startRow:6, isRange:false}` |
| Range | `API_Output!$G$5:$AK$32` | `{sheet:'API_Output', startCol:'G', startRow:5, endCol:'AK', endRow:32, isRange:true}` |

### Cell Reading API

```javascript
// Get formatted display value (calculated result for formulas)
model.getFormattedCellValue(sheetIndex, row, col) // Returns string

// Get raw content (formula string or literal value)
model.getCellContent(sheetIndex, row, col) // Returns string
```

**Important:** Both methods return **strings**. Type coercion required for numeric operations.

### Empty Cell Handling

| Cell State | getFormattedCellValue() | getCellContent() |
|------------|------------------------|------------------|
| Empty | `""` | `""` |
| Number | `"650"` | `"650"` |
| Text | `"Loan Amount"` | `"Loan Amount"` |
| Formula | `"96.875"` (result) | `"=Pricing_Output!B27"` |

### No Known Limitations

All tested scenarios worked correctly:
- Multi-cell ranges (horizontal, vertical, table)
- Large tables (28×31 = 868 cells)
- Formulas referencing other sheets
- Mixed formula/value cells in same range

---

## File References

- **Excel File:** `DSCR_NoArrayFormulas_Testing.xlsx`
- **Exploration Scripts:** `scripts/explore_*.py`
- **Exploration Report:** `docs/xlsx-exploration-report.md`
- **IronCalc Test Script:** `scripts/test-ironcalc-ranges.js`
- **Implementation:** `src/engine/namedRanges.js`, `src/engine/calculator.js`

---

## Change Log

| Date | Change |
|------|--------|
| 2026-01-31 | Initial strategy document created |
| 2026-01-31 | Decided: TABLE always OUTPUT, no INPUT tables |
| 2026-01-31 | Decided: Single table named range (not column-per-range) |
| 2026-01-31 | Defined JSON formats for all range types |
| 2026-01-31 | Added IronCalc API findings section (Task 2) |
