// Title: Export an Excel workbook to CSV with all numeric values rounded to two decimal places using Aspose.Cells for .NET
// AI Prompts: Iterate through each worksheet, round every numeric cell to two decimals, and save each sheet as a separate CSV file with Aspose.Cells. | Create a helper method that takes input and output file paths, applies Math.Round(…,2) to numeric cells, and exports the workbook to a single CSV file. | Adjust the code to keep original cell formatting while rounding numbers before calling Workbook.Save with SaveFormat.Csv.
// Common Searches: how to round numeric cells to two decimal places before exporting to CSV using Aspose.Cells .NET | Aspose.Cells export all worksheets to individual CSV files in C# | C# save Excel workbook as CSV with two-decimal rounding for financial reports | Aspose.Cells round numbers in cells programmatically prior to CSV conversion
// Tags: round numeric cells two decimals Aspose.Cells | export worksheet to CSV Aspose.Cells .NET | save each sheet as separate CSV Aspose.Cells | financial reporting CSV export Aspose.Cells | apply Math.Round to cell values Aspose.Cells

using System;
using Aspose.Cells;

namespace FinancialCsvExport
{
    // The example loads an Excel workbook, rounds every numeric cell to two decimal places, and saves the first worksheet as a CSV file using Aspose.Cells for .NET (with guidance on exporting each worksheet separately).
    class Program
    {
        static void Main(string[] args)
        {
            // Load the source workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine the used range to limit iteration
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Loop through all used cells
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = sheet.Cells[row, col];

                        // Check if the cell contains a numeric value
                        if (cell.Type == CellValueType.IsNumeric)
                        {
                            // Round the numeric value to two decimal places
                            double rounded = Math.Round(cell.DoubleValue, 2);
                            cell.PutValue(rounded);
                        }
                    }
                }
            }

            // Save the workbook as CSV.
            // By default Aspose.Cells saves only the first worksheet to CSV.
            // If you need all sheets, loop and save each separately.
            workbook.Save("output.csv", SaveFormat.Csv);
        }
    }
}
