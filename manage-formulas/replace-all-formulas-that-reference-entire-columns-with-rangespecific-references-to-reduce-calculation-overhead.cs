// Title: Replace whole‑column references with row‑bounded ranges in Excel formulas using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that scans every worksheet, detects formulas containing whole‑column references (e.g., A:A), and rewrites them to use a range from row 1 to the sheet's last used row. | Write a method that applies a regular expression to locate column‑only references in cell formulas and substitutes them with explicit row limits derived from the worksheet's MaxDataRow property.
// Common Searches: convert A:A formulas to A1:A500 with Aspose.Cells C# | limit Excel formula range to used rows for faster calculation in .NET | C# iterate over all formulas in a workbook and change whole column references | regex replace column‑only references in Excel formulas using Aspose.Cells | optimize Excel calculation speed by bounding column references in Aspose.Cells
// Tags: Aspose.Cells column‑range optimization | C# bound formula ranges | regex based formula transformation .NET | Excel calculation performance improvement | programmatic formula editing Aspose

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The program loads an Excel workbook, determines each worksheet's last used row, uses a regex to find whole‑column references in formulas, replaces them with explicit row‑bounded ranges (e.g., A1:A{lastRow}), updates the cells, and saves the modified file.
class ReplaceColumnFormulas
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the last used row in the current sheet (1‑based index)
            int lastRowIndex = sheet.Cells.MaxDataRow + 1;
            if (lastRowIndex == 0) continue; // Skip empty sheets

            // Regular expression to find whole‑column references like A:A, B:B, AA:AA, etc.
            Regex columnRefRegex = new Regex(@"([A-Z]+):\1", RegexOptions.IgnoreCase);

            // Loop through all cells that contain formulas
            foreach (Cell cell in sheet.Cells)
            {
                if (!cell.IsFormula) continue;

                string originalFormula = cell.Formula;

                // Replace each whole‑column reference with a range‑specific reference
                string updatedFormula = columnRefRegex.Replace(originalFormula, match =>
                {
                    string column = match.Groups[1].Value.ToUpper(); // Column letters
                    // Build a range from the first row to the last used row
                    return $"{column}1:{column}{lastRowIndex}";
                });

                // If any replacement occurred, assign the new formula back to the cell
                if (!originalFormula.Equals(updatedFormula, StringComparison.Ordinal))
                {
                    cell.Formula = updatedFormula;
                }
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
