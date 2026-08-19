// Title: Aspose.Cells for .NET – Position a slicer precisely using Shape.Top & Shape.Left (pixel values) in C#
// Description: C# code that creates a workbook, adds a pivot table, inserts a slicer linked to the pivot, and positions the slicer at exact pixel coordinates (Top = 200, Left = 100) by assigning values to slicer.Shape.Top and slicer.Shape.Left. The example also shows how to lock the slicer’s location so end users cannot move it, then saves the workbook.
// Keywords: Aspose.Cells slicer position | C# set slicer top left | Shape.Top pixel coordinate | Shape.Left pixel coordinate | lock slicer location Aspose.Cells | programmatic slicer placement .NET | Excel pivot table slicer C# | dashboard layout Aspose.Cells
// Common Searches: How to set slicer top and left in Aspose.Cells C# | Specify pixel coordinates for a slicer with Aspose.Cells | Lock slicer position programmatically using Aspose.Cells | Move Excel slicer to exact location in .NET | Aspose.Cells shape top left properties for slicers
// Developer Intent: Place a slicer at exact pixel coordinates and optionally lock its position using Aspose.Cells for .NET.
// Use Cases: Align a category slicer at (100 px, 200 px) to match other dashboard elements. | Prevent users from moving slicers in a generated report by locking their position. | Create a multi‑slicer layout with precise pixel placement for a custom Excel dashboard.
// AI Prompts: Write C# code with Aspose.Cells that adds a slicer to a pivot table and sets Shape.Top to 150 and Shape.Left to 80 pixels. | Show how to lock a slicer's position after positioning it with pixel coordinates in Aspose.Cells for .NET. | Provide an example that positions three slicers at different pixel locations on the same worksheet using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

namespace SlicerPositionExample
{
    // C# code that creates a workbook, adds a pivot table, inserts a slicer linked to the pivot, and positions the slicer at exact pixel coordinates (Top = 200, Left = 100) by assigning values to slicer.Shape.Top and slicer.Shape.Left. The example also shows how to lock the slicer’s location so end users cannot move it, then saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["B1"].Value = "Amount";
            sheet.Cells["A2"].Value = "Food";
            sheet.Cells["B2"].Value = 120;
            sheet.Cells["A3"].Value = "Drink";
            sheet.Cells["B3"].Value = 80;
            sheet.Cells["A4"].Value = "Supplies";
            sheet.Cells["B4"].Value = 150;

            // Add a pivot table based on the data
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D2", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Add a slicer linked to the pivot table (field index 0 = Category)
            SlicerCollection slicers = sheet.Slicers;
            int slicerIdx = slicers.Add(pivot, 10, 5, 0); // place initially at row 10, column 5
            Slicer slicer = slicers[slicerIdx];

            // Position the slicer precisely using Shape.Top and Shape.Left (pixels)
            // Top = 200 pixels from the top of the worksheet
            // Left = 100 pixels from the left of the worksheet
            slicer.Shape.Top = 200;
            slicer.Shape.Left = 100;

            // Optionally lock the position so users cannot move it via UI
            slicer.LockedPosition = true;

            // Save the workbook
            workbook.Save("SlicerPositioned.xlsx");
        }
    }
}
