// Title: Create a slicer for a ListObject and obtain its name identifier using Aspose.Cells for .NET (C#)
// Common Searches: how to add a slicer to a ListObject in Aspose.Cells C# and get its name | Aspose.Cells .NET retrieve slicer identifier after creating slicer for a table | C# example of creating an Excel slicer with Aspose.Cells and using slicer.Name as ID | add slicer to table column using Aspose.Cells and reference its name in later API calls | Aspose.Cells API guide for slicer creation and identifier extraction
// Tags: ListObject slicer addition Aspose.Cells | slicer name retrieval Aspose.Cells C# | Excel table slicer creation Aspose.Cells .NET | use slicer.Name as identifier | Aspose.Cells slicer management example

using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;
using System;
using System.IO;

// The sample creates a new workbook, adds sample sales data, defines a ListObject named "SalesData", inserts a slicer linked to the "Category" column, captures the slicer's Name property as its identifier, prints the identifier, and saves the workbook as "SlicerDemo.xlsx".
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("North");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("South");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("East");
            sheet.Cells["B4"].PutValue(200);
            sheet.Cells["A5"].PutValue("West");
            sheet.Cells["B5"].PutValue(120);

            // Create a table (ListObject) from the data range
            int firstRow = 0;          // zero‑based index
            int firstColumn = 0;
            int totalRows = 5;
            int totalColumns = 2;
            int tableIdx = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows - 1,
                firstColumn + totalColumns - 1,
                true);

            ListObject table = sheet.ListObjects[tableIdx];
            table.DisplayName = "SalesData";   // name used when adding slicers

            // Add a slicer for the "Category" column (column index 0) of the table
            int slicerIdx = sheet.Slicers.Add(table, 0, "CategorySlicer");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Use the slicer's name as an identifier (SlicerId property is not exposed)
            string slicerId = slicer.Name;

            Console.WriteLine("Slicer Id (Name): " + slicerId);

            // Save the workbook
            string outputPath = "SlicerDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
