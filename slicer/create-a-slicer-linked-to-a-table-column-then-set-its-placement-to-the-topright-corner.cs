// Title: Add a slicer linked to a table column and position it in the top‑right corner of the worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a slicer that is bound to the first column of a ListObject and place it at row 0, column 5 (top‑right area) with Aspose.Cells. | Set the slicer's Placement property to MoveAndSize and then adjust its HeightPixel to 150 and WidthPixel to 120.
// Common Searches: Aspose.Cells C# add slicer to ListObject column | C# set slicer placement to top right of Excel sheet using Aspose.Cells | How to use MoveAndSize placement type for a slicer in Aspose.Cells | Resize slicer height and width in pixels with Aspose.Cells C# | Save workbook with slicer positioned in top‑right corner Aspose.Cells
// Tags: Aspose.Cells add slicer to ListObject column | Aspose.Cells set slicer placement top right | Aspose.Cells slicer MoveAndSize placement | Aspose.Cells resize slicer pixels | Aspose.Cells create workbook with slicer

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

// Demonstrates creating a workbook, defining a table (ListObject), adding a slicer linked to the first column, positioning it at row 0 column 5 (top‑right corner), setting Placement to MoveAndSize, adjusting its height to 150 px and width to 120 px, and saving the file as SlicerTopRight.xlsx.
class SlicerTopRightDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a table (two columns)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Add a ListObject (table) covering the data range
        int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Add a slicer linked to the first column of the table.
        // Place it at row 0, column 5 (top‑right area of the sheet).
        SlicerCollection slicers = sheet.Slicers;
        int slicerIdx = slicers.Add(table, table.ListColumns[0], 0, 5);
        Slicer slicer = slicers[slicerIdx];

        // Set the placement type (how the slicer moves/resizes with cells)
        slicer.Placement = PlacementType.MoveAndSize; // top‑right corner placement

        // Optionally adjust size
        slicer.HeightPixel = 150;
        slicer.WidthPixel = 120;

        // Save the workbook
        workbook.Save("SlicerTopRight.xlsx");
    }
}
