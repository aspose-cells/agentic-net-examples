// Title: Clone a slicer’s properties to another worksheet using Aspose.Cells for .NET (C#)
// Description: Learn how to create a workbook, add a pivot table, place a slicer, customize its caption, style, layout and shape, then duplicate those settings on a new slicer in a different worksheet. The example shows property‑by‑property copying—including caption, style, show options, column count, locked position, and shape dimensions—so the second slicer looks and behaves exactly like the first.
// Keywords: Aspose.Cells slicer copy | clone slicer .NET | duplicate slicer worksheet | C# Aspose.Cells slicer properties | copy slicer style and size | pivot table slicer replication | Aspose.Cells SlicerCollection | Excel slicer automation
// Common Searches: how to copy slicer settings in Aspose.Cells | duplicate slicer on another sheet C# | clone slicer style Aspose.Cells for .NET | copy slicer shape dimensions programmatically | Aspose.Cells copy slicer properties example
// Developer Intent: Programmatically copy all visual and behavioral properties from an existing slicer to a new slicer on a different worksheet.
// Use Cases: Maintain a consistent slicer appearance across multiple dashboard sheets in an automated Excel report. | Synchronize slicer layout and filtering controls when generating multi‑sheet workbooks with the same pivot source. | Create template‑driven reports where identical slicers are needed on several worksheets without manual formatting.
// AI Prompts: Generate C# code with Aspose.Cells that clones every property of a source slicer—including caption, style, show options, column count, locked position, and shape dimensions—to a new slicer on another worksheet. | Show how to copy slicer settings while handling null references and ensuring the target slicer is linked to the same pivot table. | Write a reusable method that accepts a source Slicer object and a destination Worksheet, creates a linked slicer, and copies all visual and behavioral properties using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

// Learn how to create a workbook, add a pivot table, place a slicer, customize its caption, style, layout and shape, then duplicate those settings on a new slicer in a different worksheet. The example shows property‑by‑property copying—including caption, style, show options, column count, locked position, and shape dimensions—so the second slicer looks and behaves exactly like the first.
class SlicerCopyDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];

            // Populate sample data for a pivot table
            sheet1.Cells["A1"].Value = "Fruit";
            sheet1.Cells["B1"].Value = "Quantity";
            sheet1.Cells["A2"].Value = "Apple";
            sheet1.Cells["B2"].Value = 10;
            sheet1.Cells["A3"].Value = "Orange";
            sheet1.Cells["B3"].Value = 20;
            sheet1.Cells["A4"].Value = "Banana";
            sheet1.Cells["B4"].Value = 30;

            // Add a pivot table on the first sheet
            int pivotIdx = sheet1.PivotTables.Add("A1:B4", "D1", "FruitPivot");
            PivotTable pivot = sheet1.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

            // Add a slicer on the first worksheet
            SlicerCollection slicersSheet1 = sheet1.Slicers;
            int slicerIdx1 = slicersSheet1.Add(pivot, "F1", "Fruit");
            Slicer slicer1 = slicersSheet1[slicerIdx1];

            // Set various properties on the original slicer
            slicer1.Caption = "Fruit Selector";
            slicer1.StyleType = SlicerStyleType.SlicerStyleDark2;
            slicer1.ShowCaption = true;
            slicer1.ShowAllItems = false;
            slicer1.NumberOfColumns = 2;
            slicer1.LockedPosition = false;

            // Shape‑related properties (position and size) using points
            slicer1.Shape.Left = 100;
            slicer1.Shape.Top = 50;
            slicer1.Shape.Width = 200;
            slicer1.Shape.Height = 120;

            // Add a second worksheet where the copied slicer will be placed
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Add a slicer on the second worksheet using the same pivot table as source
            SlicerCollection slicersSheet2 = sheet2.Slicers;
            int slicerIdx2 = slicersSheet2.Add(pivot, "F1", "Fruit");
            Slicer slicer2 = slicersSheet2[slicerIdx2];

            // Copy properties from slicer1 to slicer2
            slicer2.Caption = slicer1.Caption;
            slicer2.StyleType = slicer1.StyleType;
            slicer2.ShowCaption = slicer1.ShowCaption;
            slicer2.ShowAllItems = slicer1.ShowAllItems;
            slicer2.NumberOfColumns = slicer1.NumberOfColumns;
            slicer2.LockedPosition = slicer1.LockedPosition;

            // Copy shape properties
            slicer2.Shape.Left = slicer1.Shape.Left;
            slicer2.Shape.Top = slicer1.Shape.Top;
            slicer2.Shape.Width = slicer1.Shape.Width;
            slicer2.Shape.Height = slicer1.Shape.Height;

            // Save the workbook
            workbook.Save("SlicerCopyDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
