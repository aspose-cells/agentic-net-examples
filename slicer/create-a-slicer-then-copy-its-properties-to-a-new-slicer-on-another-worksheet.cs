// Title: Copy slicer properties to a new slicer on another worksheet with Aspose.Cells for .NET (C#)
// AI Prompts: Create a slicer linked to a pivot table and transfer its caption, style, column layout, dimensions, and locked position to a second slicer on a different worksheet using Aspose.Cells in C#. | Programmatically clone all visual and behavioral attributes of an existing slicer and apply them to a new slicer on another sheet with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells example to duplicate a slicer on another sheet in C# | how to transfer slicer caption and style between worksheets using Aspose.Cells | programmatic way to clone pivot table slicer settings in .NET | copy slicer dimensions and locked position to a new slicer with Aspose.Cells
// Tags: Aspose.Cells copy slicer properties | C# pivot table slicer cloning | apply slicer style programmatically | set slicer dimensions .NET | clone slicer locked position

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The sample creates a workbook, adds sample data and a pivot table, inserts a slicer on the first worksheet with customized caption, style, column count, size, and locked position, then adds a second worksheet, creates another slicer linked to the same pivot, and copies all the first slicer's properties to the second before saving the file.
class SlicerCopyExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet1 = workbook.Worksheets[0];
        Cells cells = sheet1.Cells;

        // Populate sample data for the pivot table
        cells["A1"].Value = "Fruit";
        cells["A2"].Value = "Apple";
        cells["A3"].Value = "Orange";
        cells["A4"].Value = "Banana";
        cells["A5"].Value = "Apple";

        cells["B1"].Value = "Quantity";
        cells["B2"].Value = 10;
        cells["B3"].Value = 20;
        cells["B4"].Value = 30;
        cells["B5"].Value = 15;

        // Add a pivot table on the first worksheet
        int pivotIdx = sheet1.PivotTables.Add("A1:B5", "D1", "FruitPivot");
        PivotTable pivot = sheet1.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer on the first worksheet and set some properties
        SlicerCollection slicersSheet1 = sheet1.Slicers;
        int slicerIdx1 = slicersSheet1.Add(pivot, "F1", "Fruit");
        Slicer slicer1 = slicersSheet1[slicerIdx1];
        slicer1.Caption = "Fruit Slicer";
        slicer1.StyleType = SlicerStyleType.SlicerStyleDark1;
        slicer1.ShowCaption = true;
        slicer1.ShowAllItems = false;
        slicer1.NumberOfColumns = 2;
        slicer1.WidthPixel = 200;
        slicer1.HeightPixel = 120;
        slicer1.LockedPosition = true;

        // Add a second worksheet
        Worksheet sheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
        sheet2.Name = "Sheet2";

        // Add a slicer on the second worksheet using the same pivot table
        SlicerCollection slicersSheet2 = sheet2.Slicers;
        int slicerIdx2 = slicersSheet2.Add(pivot, "F1", "Fruit");
        Slicer slicer2 = slicersSheet2[slicerIdx2];

        // Copy properties from slicer1 to slicer2
        slicer2.Caption = slicer1.Caption;
        slicer2.StyleType = slicer1.StyleType;
        slicer2.ShowCaption = slicer1.ShowCaption;
        slicer2.ShowAllItems = slicer1.ShowAllItems;
        slicer2.NumberOfColumns = slicer1.NumberOfColumns;
        slicer2.WidthPixel = slicer1.WidthPixel;
        slicer2.HeightPixel = slicer1.HeightPixel;
        slicer2.LockedPosition = slicer1.LockedPosition;

        // Save the workbook
        workbook.Save("SlicerCopyExample.xlsx");
    }
}
