// Title: How to replace whole‑column references with row‑bounded ranges in Excel formulas using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# replace whole column reference A:A with A1:A1000 in formulas | C# code to limit Excel formula column range to the last used row using Aspose.Cells | Optimize Excel calculation by converting column‑wide formulas to specific ranges with Aspose.Cells .NET | Use regex in Aspose.Cells to change A:A to A1:A{lastRow} across all worksheets | Batch update formulas in a workbook to use bounded ranges with Aspose.Cells C#
// Tags: Aspose.Cells replace column‑wide formula references | C# limit formula range to last data row | regex transformation of Excel column references Aspose.Cells | reduce calculation overhead by bounding formula ranges | batch process worksheets to update formulas Aspose.Cells

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

// The example loads an Excel workbook, determines each worksheet's last used row, and uses a compiled regular expression to find whole‑column references (e.g., A:A) in formula cells. It replaces each match with a row‑bounded range (e.g., A1:A{lastRow}) and writes the updated formula back, then saves the modified workbook.
class ReplaceColumnFormulas
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Regular expression to find whole‑column references like A:A, B:B, AA:AA, etc.
        Regex wholeColumnRegex = new Regex(@"([A-Z]+):\1", RegexOptions.Compiled);

        // Process each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the last used row in the worksheet (0‑based index)
            int lastDataRowIndex = sheet.Cells.MaxDataRow;
            // Convert to 1‑based row number for Excel formulas
            int lastRowNumber = lastDataRowIndex + 1;

            // Iterate through all cells that contain formulas
            foreach (Cell cell in sheet.Cells)
            {
                if (!cell.IsFormula) continue;

                string originalFormula = cell.Formula;
                string updatedFormula = wholeColumnRegex.Replace(originalFormula, match =>
                {
                    // Extract the column letters (e.g., "A", "AA")
                    string column = match.Groups[1].Value;
                    // Build a range‑specific reference: Column1:ColumnLastRow
                    return $"{column}1:{column}{lastRowNumber}";
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
