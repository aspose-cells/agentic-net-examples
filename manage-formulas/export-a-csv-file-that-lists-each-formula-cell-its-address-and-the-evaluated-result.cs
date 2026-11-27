// Title: Export Excel formula cells with their addresses and evaluated values to a CSV file using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export cells that contain formulas to CSV with cell address | How to write formula results and their A1 addresses to a CSV using Aspose.Cells | Aspose.Cells iterate worksheets and save evaluated formula values to CSV in .NET | Generate CSV list of formula cells from an XLSX workbook with Aspose.Cells C# | Export calculated formula results with sheet name to CSV using Aspose.Cells API
// Tags: export formula cells to CSV Aspose.Cells | retrieve evaluated formula values .NET | cell address extraction Aspose.Cells | calculate workbook formulas before export | CSV generation for Excel formulas C#

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The example loads an Excel workbook, forces formula calculation, scans each worksheet's used range, identifies cells that contain formulas, builds their A1-style addresses prefixed with the sheet name, captures the evaluated results, escapes them for CSV compliance, and writes all entries to a CSV file named FormulaResults.csv.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Ensure all formulas are evaluated
        workbook.CalculateFormula();

        // Prepare a StringBuilder for CSV content
        StringBuilder csvBuilder = new StringBuilder();

        // Write CSV header
        csvBuilder.AppendLine("Address,Result");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            Cells cells = sheet.Cells;

            // Determine the used range to limit iteration
            int maxRow = cells.MaxDataRow;
            int maxColumn = cells.MaxDataColumn;

            // Scan each cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Check if the cell contains a formula
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        // Build the cell address in A1 style, prefixed with sheet name
                        string address = $"{sheet.Name}!{cell.Name}";

                        // Get the evaluated result as a string (handle nulls)
                        string result = cell.Value?.ToString() ?? string.Empty;

                        // Escape double quotes in the result for CSV compliance
                        result = result.Replace("\"", "\"\"");

                        // Append the line to the CSV (quote the result to preserve commas)
                        csvBuilder.AppendLine($"{address},\"{result}\"");
                    }
                }
            }
        }

        // Write the CSV content to a file (replace with desired output path)
        File.WriteAllText("FormulaResults.csv", csvBuilder.ToString());
    }
}
