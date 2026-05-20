using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

class SlicerCopyExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet1 = workbook.Worksheets[0];
        Cells cells = sheet1.Cells;

        // Populate sample data for a pivot table
        cells["A1"].Value = "Fruit";
        cells["A2"].Value = "Apple";
        cells["A3"].Value = "Orange";
        cells["A4"].Value = "Banana";
        cells["B1"].Value = "Quantity";
        cells["B2"].Value = 10;
        cells["B3"].Value = 20;
        cells["B4"].Value = 30;

        // Add a pivot table on the first sheet
        int pivotIdx = sheet1.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet1.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

        // Add a slicer on the first sheet linked to the pivot table
        int slicerIdx1 = sheet1.Slicers.Add(pivot, "E2", "Fruit");
        Slicer slicer1 = sheet1.Slicers[slicerIdx1];

        // Set various properties on the original slicer
        slicer1.Caption = "Fruit Selector";
        slicer1.StyleType = SlicerStyleType.SlicerStyleDark2;
        slicer1.NumberOfColumns = 2;
        slicer1.WidthPixel = 200;
        slicer1.HeightPixel = 120;
        slicer1.LockedPosition = true;
        slicer1.ShowCaption = true;

        // Position the slicer using its shape
        slicer1.Shape.Left = 100;
        slicer1.Shape.Top = 50;

        // Add a second worksheet
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

        // Add a slicer on the second sheet linked to the same pivot table
        int slicerIdx2 = sheet2.Slicers.Add(pivot, "E2", "Fruit");
        Slicer slicer2 = sheet2.Slicers[slicerIdx2];

        // Copy properties from slicer1 to slicer2
        slicer2.Caption = slicer1.Caption;
        slicer2.StyleType = slicer1.StyleType;
        slicer2.NumberOfColumns = slicer1.NumberOfColumns;
        slicer2.WidthPixel = slicer1.WidthPixel;
        slicer2.HeightPixel = slicer1.HeightPixel;
        slicer2.LockedPosition = slicer1.LockedPosition;
        slicer2.ShowCaption = slicer1.ShowCaption;

        // Copy shape positioning
        slicer2.Shape.Left = slicer1.Shape.Left;
        slicer2.Shape.Top = slicer1.Shape.Top;

        // Save the workbook
        workbook.Save("SlicerCopyExample.xlsx");
    }
}