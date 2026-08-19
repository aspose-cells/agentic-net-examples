// Title: Aspose.Cells for .NET – Add a slicer linked to a table column and position it in the top‑right corner (C#)
// Description: Demonstrates how to create a workbook, define a ListObject (table), insert a slicer linked to the first column, set its title, configure Placement to MoveAndSize, and place the slicer at the worksheet’s top‑right corner using TopPixel and LeftPixel coordinates before saving the file.
// Keywords: Aspose.Cells slicer C# | add slicer to ListObject | slicer placement top right | PlacementType MoveAndSize | TopPixel LeftPixel Aspose.Cells | C# Excel slicer example | Aspose.Cells table filter
// Common Searches: Aspose.Cells add slicer to table column C# | position slicer at top right of worksheet Aspose.Cells | set slicer placement MoveAndSize Aspose.Cells .NET | C# code for slicer pixel coordinates Aspose.Cells | how to anchor slicer to top‑right corner in Excel using Aspose
// Developer Intent: Insert a slicer linked to a table column and anchor it at the worksheet’s top‑right corner.
// Use Cases: Create an interactive filter for a dashboard where the slicer is always visible in the top‑right area. | Automate report generation with consistent slicer placement across multiple sheets. | Adjust slicer location programmatically when sheet dimensions change to maintain a fixed top‑right position.
// AI Prompts: Generate C# code with Aspose.Cells that adds a slicer for the first column of a ListObject, sets Placement to MoveAndSize, and positions it at the top‑right corner using pixel coordinates. | Show how to configure slicer size (WidthPixel, HeightPixel) and location (TopPixel, LeftPixel) after linking it to a table column in Aspose.Cells for .NET. | Provide a complete Aspose.Cells example that creates sample data, defines a table, adds a linked slicer, moves it to the top‑right of the sheet, and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, define a ListObject (table), insert a slicer linked to the first column, set its title, configure Placement to MoveAndSize, and place the slicer at the worksheet’s top‑right corner using TopPixel and LeftPixel coordinates before saving the file.
class SlicerTopRightDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the table (two columns)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["A5"].PutValue("A");
        sheet.Cells["B5"].PutValue(40);
        sheet.Cells["A6"].PutValue("B");
        sheet.Cells["B6"].PutValue(50);

        // Add a ListObject (table) that covers the data range
        int tableIndex = sheet.ListObjects.Add("A1", "B6", true);
        ListObject table = sheet.ListObjects[tableIndex];
        table.TableStyleType = TableStyleType.TableStyleMedium2;

        // Add a slicer linked to the first column ("Category") of the table
        // Place the slicer starting at row 8, column 8 (zero‑based indices)
        SlicerCollection slicers = sheet.Slicers;
        int slicerIndex = slicers.Add(table, table.ListColumns[0], 8, 8);
        Slicer slicer = slicers[slicerIndex];

        // Set slicer properties
        slicer.Title = "Category Filter";

        // Set the placement type (how the slicer moves/resizes with cells)
        slicer.Placement = PlacementType.MoveAndSize;

        // Position the slicer at the top‑right corner of the worksheet
        // TopPixel = 0 places it at the top edge
        // LeftPixel is set far to the right; here we use a large value (e.g., 800 pixels)
        slicer.TopPixel = 0;
        slicer.LeftPixel = 800;

        // Optionally adjust size
        slicer.WidthPixel = 150;
        slicer.HeightPixel = 200;

        // Save the workbook
        workbook.Save("SlicerTopRightDemo.xlsx");
    }
}
