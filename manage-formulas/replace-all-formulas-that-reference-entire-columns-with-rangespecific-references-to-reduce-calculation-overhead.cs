// Title: Aspose.Cells C# – Replace Whole‑Column References with Specific Ranges to Speed Up Formula Calculation
// Description: Loads a workbook, disables the calculation chain, scans all worksheets for formulas, uses a regex to locate whole‑column references (e.g., A:A), determines the last populated row in each column, replaces the reference with a precise range (A1:A{lastRow}), recalculates, and saves the file—significantly reducing calculation overhead.
// Keywords: Aspose.Cells | C# | replace whole column reference | Excel formula optimization | dynamic range replacement | GetLastDataRow | disable calculation chain | performance | Excel workbook | formula calculation overhead
// Common Searches: Aspose.Cells replace A:A with A1:A100 | C# code to convert whole column references to ranges | How to speed up Excel formula calculation using Aspose.Cells | Disable calculation chain Aspose.Cells bulk formula update | Find and replace column references in Excel formulas .NET
// Developer Intent: Automatically convert every whole‑column reference in an Excel workbook to the smallest data‑containing range, reducing calculation load.
// Use Cases: Transform SUM(A:A) to SUM(A1:A{lastRow}) after data import | Prepare large financial models for faster recalculation | Batch‑update formulas in templates before publishing | Optimize performance of dashboards that reference entire columns
// AI Prompts: Write C# using Aspose.Cells that scans all worksheets, finds whole‑column references in formulas, and replaces them with a range from row 1 to the last non‑empty row. | Explain the impact of disabling the calculation chain during bulk formula edits and how to re‑enable it safely. | Provide sample code to log each formula change (original vs new) to a CSV file while processing the workbook.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Loads a workbook, disables the calculation chain, scans all worksheets for formulas, uses a regex to locate whole‑column references (e.g., A:A), determines the last populated row in each column, replaces the reference with a precise range (A1:A{lastRow}), recalculates, and saves the file—significantly reducing calculation overhead.
    class ReplaceColumnFormulas
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Disable calculation chain for faster processing (optional)
            workbook.Settings.FormulaSettings.EnableCalculationChain = false;

            // Regex to match whole‑column references like A:A, $A:$A, B:B etc.
            Regex columnRefRegex = new Regex(@"\$?([A-Z]+):\$?\1", RegexOptions.Compiled);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Cells cells = sheet.Cells;

                // Enumerate all cells that contain formulas
                foreach (Cell cell in cells)
                {
                    if (!cell.IsFormula) continue;

                    string originalFormula = cell.Formula;
                    string updatedFormula = columnRefRegex.Replace(originalFormula, match =>
                    {
                        // Extract column letters (e.g., "A")
                        string colLetters = match.Groups[1].Value;

                        // Convert column letters to zero‑based index
                        int colIndex = CellsHelper.ColumnNameToIndex(colLetters);

                        // Determine the last row that actually contains data in this column
                        int lastDataRow = cells.GetLastDataRow(colIndex);
                        // If the column is empty, default to row 1 to avoid invalid range
                        if (lastDataRow < 0) lastDataRow = 1;

                        // Build a range that starts at row 1 and ends at the last data row
                        string newRange = $"{colLetters}1:{colLetters}{lastDataRow + 1}"; // +1 because GetLastDataRow is zero‑based
                        return newRange;
                    });

                    // If the formula was changed, assign the new formula back to the cell
                    if (!originalFormula.Equals(updatedFormula, StringComparison.Ordinal))
                    {
                        cell.Formula = updatedFormula;
                    }
                }
            }

            // Recalculate all formulas after modifications
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
