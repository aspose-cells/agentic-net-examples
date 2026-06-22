using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate worksheet with sample data for the pivot table
        sheet.Cells["A1"].Value = "Category";
        sheet.Cells["B1"].Value = "Amount";
        sheet.Cells["A2"].Value = "Fruit";
        sheet.Cells["B2"].Value = 100;
        sheet.Cells["A3"].Value = "Vegetable";
        sheet.Cells["B3"].Value = 150;
        sheet.Cells["A4"].Value = "Fruit";
        sheet.Cells["B4"].Value = 200;

        // Add a pivot table that uses the data range A1:B4 and place it at D2
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D2", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Configure the pivot table: Category as row field, Amount as data field
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Add a slicer linked to the pivot table for the "Category" field.
        // The slicer will be placed with its top‑left corner at cell F2.
        int slicerIndex = sheet.Slicers.Add(pivot, "F2", "Category");
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Freeze the slicer’s position so users cannot move or resize it.
        slicer.LockedPosition = true;

        // Save the workbook to a file.
        workbook.Save("SlicerPivotFreeze.xlsx");
    }
}