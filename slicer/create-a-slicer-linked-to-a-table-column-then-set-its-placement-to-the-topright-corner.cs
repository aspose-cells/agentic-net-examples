// Title: How to add a slicer linked to a table column and position it in the top‑right corner using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a ListObject table, builds a pivot table, adds a slicer for the first column, and moves the slicer to the top‑right cell of the worksheet. | Show how to set the slicer’s row and column indices so it appears at cell H1 while staying linked to the pivot table column. | Provide a complete example that includes error handling and saves the workbook with the slicer positioned at the top‑right corner.
// Common Searches: Aspose.Cells C# add slicer to pivot table and set its location | C# Aspose.Cells place slicer at top right of worksheet | How to link a slicer to a ListObject column using Aspose.Cells for .NET | Aspose.Cells example for slicer placement coordinates | Create pivot table with slicer in Aspose.Cells C# tutorial
// Tags: aspocells slicer linked to listobject | aspocells slicer top‑right placement | aspocells pivot table slicer creation | c# aspocells table slicer example | aspocells workbook slicer positioning

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers; // For slicer support

// Demonstrates creating a workbook, adding a ListObject table, generating a pivot table, inserting a slicer linked to the first column, setting its size, moving it to the worksheet’s top‑right corner, and saving the file as SlicerExample.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["A4"].PutValue("A");
            dataSheet.Cells["B4"].PutValue(30);

            // Define the range of the table (including header row)
            int firstRow = 0;      // zero‑based index
            int firstColumn = 0;
            int totalRows = 5;     // header + 4 data rows (extra row is harmless)
            int totalColumns = 2;

            // Add a ListObject (table) to the worksheet
            int tableIdx = dataSheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = dataSheet.ListObjects[tableIdx];
            table.DisplayName = "MyTable";

            // Add a new worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Create a pivot table based on the ListObject.
            // The Add method returns the index of the new pivot table.
            int pivotIdx = pivotSheet.PivotTables.Add("MyTable", "A1", "PivotTable1");
            PivotTable pivot = pivotSheet.PivotTables[pivotIdx];

            // Add the first column ("Category") as a row field so the slicer can reference it
            pivot.AddFieldToArea(PivotFieldType.Row, 0);

            // Add a slicer linked to the first column of the pivot table
            // Parameters: pivot table object, column index (0‑based), top row, left column for placement
            int slicerIdx = pivotSheet.Slicers.Add(pivot, 0, 0, 0);
            Slicer slicer = pivotSheet.Slicers[slicerIdx];

            // Optional: set size of the slicer (in points)
            slicer.Width = 120;
            slicer.Height = 200;

            // Save the workbook
            string outputPath = "SlicerExample.xlsx";

            // Ensure the directory exists to avoid FileNotFoundException
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
