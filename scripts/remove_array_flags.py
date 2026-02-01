import openpyxl
from openpyxl.worksheet.formula import ArrayFormula
import sys
import os

input_file = sys.argv[1] if len(sys.argv) > 1 else "../DSCR_Complete_Pricing_Engine_Dev_NoArrayTesting.xlsx"
output_file = sys.argv[2] if len(sys.argv) > 2 else "../DSCR_NoArrayFormulas_STYLED.xlsx"

print(f"\n=== Remove CSE Array Flags (with styling preserved) ===")
print(f"Input:  {input_file}")
print(f"Output: {output_file}\n")

# Load workbook with all formatting preserved
wb = openpyxl.load_workbook(input_file)

fixed_count = 0
fixed_cells = []

# Known array formula cells from our analysis
array_cells = [
    ("LTV_Matrix", "E48"),
    ("LTV_Matrix", "E50"),
    ("LTV_Matrix", "E52"),
    ("Adjustments", "C40"),
    ("Adjustments", "D45"),
    ("Adjustments", "D46"),
    ("Adjustments", "D47"),
    ("Adjustments", "D48"),
    ("Adjustments", "D49"),
    ("Adjustments", "D50"),
    ("Adjustments", "D51"),
    ("Adjustments", "D52"),
    ("Adjustments", "D53"),
    ("Adjustments", "D54"),
    ("Adjustments", "D55"),
    ("Adjustments", "D56"),
    ("Adjustments", "D57"),
    ("Adjustments", "D58"),
    ("Adjustments", "D63"),
    ("Adjustments", "D64"),
    ("Adjustments", "D65"),
    ("Adjustments", "D66"),
    ("Adjustments", "D67"),
    ("Adjustments", "D68"),
    ("Adjustments", "D69"),
    ("Adjustments", "D88"),
    ("Pricing_Output", "W29"),
    ("Pricing_Output", "W31"),
    ("Pricing_Output", "W33"),
]

for sheet_name, cell_addr in array_cells:
    ws = wb[sheet_name]
    cell = ws[cell_addr]

    # Check if it's an ArrayFormula
    if isinstance(cell.value, ArrayFormula):
        formula_text = cell.value.text
        # Convert to regular formula
        cell.value = f"={formula_text}"
        fixed_count += 1
        fixed_cells.append(f"{sheet_name}!{cell_addr}")
        print(f"  Fixed: {sheet_name}!{cell_addr}")
    elif cell.value and str(cell.value).startswith("="):
        # Already a regular formula, check if it has array_formula attribute
        print(f"  Already regular: {sheet_name}!{cell_addr}")
    else:
        print(f"  Checking: {sheet_name}!{cell_addr} - Value type: {type(cell.value)}")

# Save to new file
wb.save(output_file)

print(f"\n[OK] Fixed {fixed_count} array formulas")
print(f"[OK] Saved to: {output_file}")
print(f"\nStyling and formatting preserved!")
