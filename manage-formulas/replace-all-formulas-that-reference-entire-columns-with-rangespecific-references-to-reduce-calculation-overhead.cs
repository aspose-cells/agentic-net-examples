using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

class ReplaceColumnFormulas
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook wb = new Workbook("input.xlsx");

        // Process each worksheet in the workbook
        foreach (Worksheet ws in wb.Worksheets)
        {
            // Determine the last row that contains data in the current sheet
            int lastDataRowZeroBased = ws.Cells.MaxDataRow; // zero‑based index
            if (lastDataRowZeroBased < 0) lastDataRowZeroBased = 0; // handle empty sheets
            int lastDataRow = lastDataRowZeroBased + 1; // convert to 1‑based row number for Excel addresses

            // Iterate through all cells that contain formulas
            foreach (Cell cell in ws.Cells)
            {
                if (!cell.IsFormula) continue;

                string originalFormula = cell.Formula;

                // Regex to find whole‑column references like A:A, $A:$A, Sheet1!B:B, etc.
                // It matches two column identifiers separated by a colon with no row numbers.
                string pattern = @"(\$?)([A-Z]+)(?::)(\$?)([A-Z]+)";

                string updatedFormula = Regex.Replace(
                    originalFormula,
                    pattern,
                    match =>
                    {
                        // Extract column parts
                        string colPrefix1 = match.Groups[1].Value; // optional $
                        string col1 = match.Groups[2].Value;      // column letters
                        string colPrefix2 = match.Groups[3].Value; // optional $
                        string col2 = match.Groups[4].Value;      // column letters

                        // Ensure both sides refer to the same column (ignore $)
                        if (!string.Equals(col1, col2, StringComparison.OrdinalIgnoreCase))
                            return match.Value; // not a whole‑column reference, leave unchanged

                        // Build a range that spans from row 1 to the last data row
                        string startRef = $"{colPrefix1}{col1}$1";
                        string endRef   = $"{colPrefix2}{col2}${lastDataRow}";
                        return $"{startRef}:{endRef}";
                    },
                    RegexOptions.IgnoreCase);

                // If the formula was changed, set the new formula back to the cell
                if (!string.Equals(originalFormula, updatedFormula, StringComparison.Ordinal))
                {
                    cell.SetFormula(updatedFormula, new FormulaParseOptions());
                }
            }
        }

        // Recalculate all formulas after the replacements
        wb.CalculateFormula();

        // Save the modified workbook
        wb.Save("output.xlsx");
    }
}