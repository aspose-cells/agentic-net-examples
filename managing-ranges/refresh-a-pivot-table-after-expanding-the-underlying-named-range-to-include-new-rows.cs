// Title: Expand a named range and refresh all pivot tables in an Excel workbook with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that enlarges an existing named range by a given number of rows and then calls RefreshData and CalculateData on every pivot table in the workbook using Aspose.Cells. | Show how to modify the RefersTo property of a named range and programmatically update all dependent pivot tables in a .xlsx file with Aspose.Cells. | Provide a step‑by‑step example that adds rows to a data source, updates the named range address, and refreshes pivot caches in C#.
// Common Searches: aspnet c# expand named range and refresh pivot tables using Aspose.Cells | how to programmatically change RefersTo of a named range in Aspose.Cells .NET | refresh pivot cache after adding rows to data source with Aspose.Cells C# | update all pivot tables after extending data range in Excel via Aspose.Cells | Aspose.Cells example for dynamic pivot data source expansion
// Tags: named range expansion Aspose.Cells | pivot table refresh Aspose.Cells | update RefersTo property C# | refresh pivot caches .xlsx Aspose.Cells | dynamic data source for pivot tables .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example loads a workbook, retrieves a named range, expands its RefersTo address to include additional rows, iterates through every worksheet to refresh and recalculate each pivot table, and saves the modified file.
class PivotTableRefreshExample
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the named range that serves as the data source for the pivot table
            // Replace "DataRange" with the actual name of your named range
            Name dataRange = workbook.Worksheets.Names["DataRange"];
            if (dataRange == null)
            {
                Console.WriteLine("Error: Named range \"DataRange\" not found.");
                return;
            }

            // Determine the new size of the range (e.g., add 10 more rows)
            int additionalRows = 10;

            // Obtain the current range represented by the named range
            Aspose.Cells.Range range = dataRange.GetRange();

            // Calculate new last row index
            int startRow = range.FirstRow;
            int startColumn = range.FirstColumn;
            int rowCount = range.RowCount;
            int columnCount = range.ColumnCount;
            int newRowCount = rowCount + additionalRows;
            int newEndRow = startRow + newRowCount - 1;
            int endColumn = startColumn + columnCount - 1;

            // Build the new address string (e.g., Sheet1!A1:C30)
            string sheetName = range.Worksheet.Name;
            string startCell = CellsHelper.CellIndexToName(startRow, startColumn);
            string endCell = CellsHelper.CellIndexToName(newEndRow, endColumn);
            dataRange.RefersTo = $"'{sheetName}'!{startCell}:{endCell}";

            // Refresh all pivot tables that use this named range
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (PivotTable pivot in sheet.PivotTables)
                {
                    // Refresh the pivot cache data and recalculate
                    pivot.RefreshData();
                    pivot.CalculateData();
                }
            }

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
