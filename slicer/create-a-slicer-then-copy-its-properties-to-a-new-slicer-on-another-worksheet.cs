using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

class SlicerCopyDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet1 = workbook.Worksheets[0];
        Cells cells = sheet1.Cells;

        // Populate sample data for a pivot table
        cells["A1"].Value = "Fruit";
        cells["B1"].Value = "Year";
        cells["C1"].Value = "Amount";

        string[] fruits = { "Apple", "Banana", "Apple", "Banana", "Apple", "Banana" };
        int[] years = { 2020, 2020, 2021, 2021, 2022, 2022 };
        int[] amounts = { 50, 70, 60, 80, 55, 85 };

        for (int i = 0; i < fruits.Length; i++)
        {
            cells[i + 1, 0].Value = fruits[i];
            cells[i + 1, 1].Value = years[i];
            cells[i + 1, 2].Value = amounts[i];
        }

        // Add a pivot table on Sheet1
        PivotTableCollection pivots = sheet1.PivotTables;
        int pivotIdx = pivots.Add("=Sheet1!A1:C7", "E3", "FruitPivot");
        PivotTable pivot = pivots[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Column, "Year");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer on Sheet1 linked to the pivot table
        SlicerCollection slicers1 = sheet1.Slicers;
        int slicerIdx1 = slicers1.Add(pivot, "G3", "Fruit");
        Slicer slicer1 = slicers1[slicerIdx1];

        // Set various properties on the original slicer
        slicer1.Caption = "Fruit Selector";
        slicer1.StyleType = SlicerStyleType.SlicerStyleLight2;
        slicer1.LockedPosition = false;
        slicer1.ShowCaption = true;
        slicer1.NumberOfColumns = 2;
        slicer1.ColumnWidth = 80;
        slicer1.RowHeight = 20;
        slicer1.WidthPixel = 200;
        slicer1.HeightPixel = 120;
        slicer1.LeftPixel = 10;
        slicer1.TopPixel = 10;

        // Add a second worksheet where the new slicer will be placed
        Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

        // Add a slicer on Sheet2 using the same pivot table as data source
        SlicerCollection slicers2 = sheet2.Slicers;
        int slicerIdx2 = slicers2.Add(pivot, "A1", "Fruit");
        Slicer slicer2 = slicers2[slicerIdx2];

        // Copy properties from slicer1 to slicer2
        slicer2.Caption = slicer1.Caption;
        slicer2.StyleType = slicer1.StyleType;
        slicer2.LockedPosition = slicer1.LockedPosition;
        slicer2.ShowCaption = slicer1.ShowCaption;
        slicer2.NumberOfColumns = slicer1.NumberOfColumns;
        slicer2.ColumnWidth = slicer1.ColumnWidth;
        slicer2.RowHeight = slicer1.RowHeight;
        slicer2.WidthPixel = slicer1.WidthPixel;
        slicer2.HeightPixel = slicer1.HeightPixel;
        slicer2.LeftPixel = slicer1.LeftPixel;
        slicer2.TopPixel = slicer1.TopPixel;

        // Additionally copy shape-level placement if needed
        slicer2.Shape.Placement = slicer1.Shape.Placement;

        // Save the workbook
        workbook.Save("SlicerCopyDemo.xlsx");
    }
}