// Title: C# – Replace Whole‑Column References with Exact Ranges in Excel using Aspose.Cells
// Description: This example loads an Excel workbook, scans every worksheet for formulas that contain whole‑column references (e.g., A:A), determines the last populated row for each column, rewrites the formula to a precise range such as A1:A150, recalculates the workbook, and saves the result. Limiting the range reduces calculation time and memory usage in large spreadsheets.
// Keywords: Aspose.Cells C# | Excel formula optimization | replace whole column reference | range specific reference | GetLastDataRow Aspose | regex column detection | performance boost Excel .NET | calculate formula Aspose.Cells | large workbook speed | US developers | UK developers | India developers
// Common Searches: how to change A:A to A1:A100 with Aspose.Cells | optimize Excel formulas by limiting ranges in C# | Aspose.Cells replace entire column references | get last data row for a column Aspose.Cells | reduce calculation overhead whole column Excel .NET
// Developer Intent: Convert all whole‑column references in a workbook to the smallest data‑driven range to improve calculation performance.
// Use Cases: Transform =SUM(A:A) into =SUM(A1:A{lastRow}) before publishing a financial model. | Update data‑validation or chart source formulas that use full columns to explicit ranges for faster refresh. | Automate workbook cleanup in a CI pipeline to ensure optimal performance on client machines.
// AI Prompts: Write C# code with Aspose.Cells that finds and replaces any whole‑column reference in formulas with a range from row 1 to the column's last data row. | Explain the behavior of Worksheet.Cells.GetLastDataRow and how to handle empty columns when adjusting formulas. | Suggest alternative strategies for minimizing formula calculation time in massive Excel files processed with Aspose.Cells.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example loads an Excel workbook, scans every worksheet for formulas that contain whole‑column references (e.g., A:A), determines the last populated row for each column, rewrites the formula to a precise range such as A1:A150, recalculates the workbook, and saves the result. Limiting the range reduces calculation time and memory usage in large spreadsheets.
    class ReplaceWholeColumnReferences
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Regular expression to find whole‑column references like A:A, B:B, etc.
            Regex wholeColumnRegex = new Regex(@"\b([A-Z]+):\1\b", RegexOptions.Compiled);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the maximum used row and column in the sheet
                int maxRow = sheet.Cells.MaxDataRow;      // zero‑based index of the last row with data
                int maxCol = sheet.Cells.MaxDataColumn;   // zero‑based index of the last column with data

                // Scan every cell that may contain a formula
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell == null || !cell.IsFormula) continue;

                        string originalFormula = cell.Formula;
                        if (string.IsNullOrEmpty(originalFormula)) continue;

                        // Replace each whole‑column reference with a range‑specific reference
                        string updatedFormula = wholeColumnRegex.Replace(originalFormula, match =>
                        {
                            string colLetter = match.Groups[1].Value; // e.g. "A"

                            // Convert column letter to zero‑based column index
                            int colIndex = 0;
                            foreach (char ch in colLetter)
                            {
                                colIndex = colIndex * 26 + (ch - 'A' + 1);
                            }
                            colIndex--; // zero‑based

                            // Get the last row that actually contains data in this column
                            int lastDataRow = sheet.Cells.GetLastDataRow(colIndex);
                            // If the column is empty, fall back to the sheet's last used row
                            if (lastDataRow < 0) lastDataRow = maxRow;

                            // Build a range like A1:A{lastRow+1}
                            string range = $"{colLetter}1:{colLetter}{lastDataRow + 1}";
                            return range;
                        });

                        // If the formula changed, update the cell
                        if (!originalFormula.Equals(updatedFormula, StringComparison.Ordinal))
                        {
                            // Preserve parsing options (default options are sufficient here)
                            cell.SetFormula(updatedFormula, new FormulaParseOptions());
                        }
                    }
                }
            }

            // Recalculate all formulas after the modifications (optional but recommended)
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}
