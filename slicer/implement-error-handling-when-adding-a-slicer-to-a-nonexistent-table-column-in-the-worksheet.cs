// Title: How to catch and log exceptions when adding a slicer to a non‑existent column in Aspose.Cells for .NET
// AI Prompts: Generate C# code that adds a slicer to a pivot table with Aspose.Cells and wraps the worksheet.Slicers.Add call in a try‑catch block to handle invalid column indexes. | Show how to verify that a column index exists in a ListObject before creating a slicer using Aspose.Cells in C#. | Create a sample that logs the exception message when worksheet.Slicers.Add throws an error because the specified table column is missing. | Provide a reusable pattern for global error handling around slicer creation in an Aspose.Cells workbook.
// Common Searches: aspocells c# slicer add invalid column index exception handling | how to validate slicer column before adding in Aspose.Cells .NET | catch error when adding slicer to pivot table Aspose.Cells | log slicer creation failure Aspose.Cells workbook C#
// Tags: aspocells slicer column validation | c# exception handling aspocells slicer | add slicer to pivot table aspocells | invalid column index error aspocells | logging slicer creation errors c#

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;   // For ListObject
using Aspose.Cells.Pivot;    // For PivotTable

// Demonstrates creating a workbook with a table and pivot table, then attempts to add a slicer using an out‑of‑range column index, catches the resulting exception, logs the error message, and saves the workbook.
class SlicerErrorHandling
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "DataSheet";

            // Populate sample data for a table
            worksheet.Cells["A1"].PutValue("ID");
            worksheet.Cells["B1"].PutValue("Name");
            worksheet.Cells["A2"].PutValue(1);
            worksheet.Cells["B2"].PutValue("Alice");
            worksheet.Cells["A3"].PutValue(2);
            worksheet.Cells["B3"].PutValue("Bob");

            // Define the range of the table (including header row)
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int totalRows = 3;     // rows 1‑3 (header + 2 data rows)
            int totalColumns = 2;  // columns A‑B

            // Add a ListObject (table) to the worksheet
            int tableIndex = worksheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = worksheet.ListObjects[tableIndex];
            table.DisplayName = "Employees";

            // Build the source range string for the pivot table (including header)
            string startCell = CellsHelper.CellIndexToName(firstRow, firstColumn);
            string endCell = CellsHelper.CellIndexToName(firstRow + totalRows - 1, firstColumn + totalColumns - 1);
            string sourceData = $"'{worksheet.Name}'!{startCell}:{endCell}";

            // Create a PivotTable based on the table (required for slicer API in this version)
            int pivotIndex = worksheet.PivotTables.Add(sourceData, "E1", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Attempt to add a slicer for a column that does NOT exist (using an invalid column index)
            try
            {
                // Column index 5 does not exist in the pivot table; this will raise an exception
                worksheet.Slicers.Add(pivotTable, 5, 5, 0);
            }
            catch (Exception ex)
            {
                // Handle the error gracefully – log or inform the user
                Console.WriteLine($"Error adding slicer: {ex.Message}");
            }

            // Save the workbook
            workbook.Save("SlicerErrorHandling.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
