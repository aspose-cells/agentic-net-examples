// Title: How to add a slicer linked to a table column in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add slicer to ListObject column for interactive filtering | Create Excel slicer programmatically with Aspose.Cells .NET example | Set slicer position and size in an Aspose.Cells workbook using C# | Link a slicer to a table column and save as .xlsx with Aspose.Cells | Aspose.Cells .NET code to customize slicer caption and dimensions
// Tags: create slicer for Excel table Aspose.Cells | configure slicer size and caption .NET | link slicer to table column Aspose.Cells API | programmatic interactive filtering Aspose.Cells | add ListObject slicer C# Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

// The example demonstrates how to create a new workbook, populate a two‑column table, add a ListObject named "MyTable", insert a slicer linked to the "Category" column at cell D2, customize its name, caption, width, and height, and finally save the workbook as SlicerDemo.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("C");
            sheet.Cells["B5"].PutValue(40);

            // Define the range for the table (including header)
            int firstRow = 0;      // zero‑based index for row 1
            int firstColumn = 0;   // zero‑based index for column A
            int totalRows = 5;     // rows 1‑5
            int totalColumns = 2;  // columns A‑B

            // Add a table (ListObject) to the worksheet
            int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];
            // Use DisplayName to set a friendly name for the table
            table.DisplayName = "MyTable";

            // Add a slicer linked to the "Category" column (column index 0 in the table)
            // Place the slicer at cell D2
            int slicerIndex = sheet.Slicers.Add(table, 0, "D2");
            Slicer slicer = sheet.Slicers[slicerIndex];
            slicer.Name = "CategorySlicer";
            slicer.Caption = "Category";
            // Set slicer size (points)
            slicer.Width = 150;
            slicer.Height = 200;

            // Save the workbook to a file
            string outputPath = "SlicerDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
