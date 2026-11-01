// Title: Apply a custom number format to a PivotTable's grand total row with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells how to format the grand total row of a pivot table | apply custom number format to pivot table totals using Aspose.Cells .NET | set red negative number format for pivot table grand total in Aspose.Cells | format numeric cells in the grand total row of an Excel pivot table with Aspose.Cells | Aspose.Cells change number format of pivot table grand total row programmatically
// Tags: pivot table grand total numeric style Aspose.Cells | Aspose.Cells set custom numeric style C# | format numeric cells in pivot table totals | apply style to pivot table grand total row | Aspose.Cells custom style for Excel pivot totals

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads an Excel workbook, locates the first PivotTable, determines the grand total row, and applies a custom number format ("#,##0.00;[Red]-#,##0.00") to each numeric cell in that row using Aspose.Cells before saving the file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook that contains a PivotTable
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (modify index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one PivotTable
            if (sheet.PivotTables.Count > 0)
            {
                // Get the first PivotTable in the worksheet
                PivotTable pivot = sheet.PivotTables[0];

                // Retrieve the range occupied by the PivotTable (TableRange1 is the correct property)
                CellArea range = pivot.TableRange1;

                // Starting position of the PivotTable
                int startRow = range.StartRow;
                int startCol = range.StartColumn;

                // Total size of the PivotTable
                int totalRows = range.EndRow - range.StartRow + 1;
                int totalCols = range.EndColumn - range.StartColumn + 1;

                // Row index of the Grand Total row (last row of the PivotTable)
                int grandTotalRow = startRow + totalRows - 1;

                // Define the custom number format for totals
                const string customFormat = "#,##0.00;[Red]-#,##0.00";

                // Apply the custom format to each numeric cell in the Grand Total row
                for (int col = startCol; col < startCol + totalCols; col++)
                {
                    Cell cell = sheet.Cells[grandTotalRow, col];

                    // Apply only to cells that contain numeric values
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        Style style = cell.GetStyle();
                        style.Custom = customFormat;
                        cell.SetStyle(style);
                    }
                }
            }

            // Save the workbook with the applied formatting
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
