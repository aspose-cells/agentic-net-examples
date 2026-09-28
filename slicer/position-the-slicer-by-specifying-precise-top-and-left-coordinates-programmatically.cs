// Title: Set exact pixel coordinates for a slicer’s position and lock it using Aspose.Cells for .NET
// AI Prompts: Assign specific pixel values to a slicer’s Shape.Left and Shape.Top properties in a C# workbook with Aspose.Cells. | Prevent users from moving a slicer after setting its pixel coordinates in an Aspose.Cells-generated worksheet.
// Common Searches: Aspose.Cells C# how to move a slicer to a specific left and top pixel location | programmatically lock slicer position in an Excel workbook using Aspose.Cells .NET | set slicer shape coordinates in pixels with Aspose.Cells for .NET example
// Tags: set slicer shape left top Aspose.Cells | slicer fixed location .NET | position slicer pixel coordinates Excel | Aspose.Cells slicer placement example

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Creates a workbook, adds sample data and a pivot table, inserts a slicer linked to the pivot, sets the slicer’s Shape.Left to 150 and Shape.Top to 80 pixels, locks its position, and saves the file as SlicerPositioned.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Fruit");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(200);

        // Add a pivot table
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Add a slicer linked to the pivot table
        int slicerIdx = sheet.Slicers.Add(pivot, "F1", "Fruit");
        Slicer slicer = sheet.Slicers[slicerIdx];

        // Position the slicer precisely using its Shape's pixel coordinates
        // Example: 150 pixels from the left edge and 80 pixels from the top edge of the worksheet
        slicer.Shape.Left = 150; // Horizontal offset in pixels
        slicer.Shape.Top = 80;   // Vertical offset in pixels

        // Optionally lock the slicer position so it cannot be moved by the user
        slicer.LockedPosition = true;

        // Save the workbook
        workbook.Save("SlicerPositioned.xlsx");
    }
}
