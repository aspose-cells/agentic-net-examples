using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data
        cells["A1"].PutValue("Fruit");
        cells["B1"].PutValue("Year");
        cells["C1"].PutValue("Amount");

        string[] fruits = { "Apple", "Orange", "Apple", "Orange", "Apple", "Orange" };
        int[] years = { 2020, 2020, 2021, 2021, 2022, 2022 };
        int[] amounts = { 50, 70, 60, 80, 55, 85 };

        for (int i = 0; i < fruits.Length; i++)
        {
            cells[i + 1, 0].PutValue(fruits[i]);
            cells[i + 1, 1].PutValue(years[i]);
            cells[i + 1, 2].PutValue(amounts[i]);
        }

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:C7", "E1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Column, "Year");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
        pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer for the "Fruit" field using a destination cell name
        int slicerIndex = sheet.Slicers.Add(pivot, "G1", "Fruit");
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Configure slicer appearance and behavior
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
        slicer.Caption = "Fruit Filter";
        slicer.LockedPosition = false;
        slicer.Placement = PlacementType.FreeFloating; // Allows free movement

        // Render the workbook (including the slicer) to PDF
        workbook.Save("SlicerDemo.pdf", SaveFormat.Pdf);
    }
}