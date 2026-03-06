using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Add a pivot table (source range A1:B4, placed at D1)
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];

        // Configure the pivot table fields
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

        // Add a slicer linked to the pivot table, positioned at cell F1, filtering by the "Fruit" field
        int slicerIndex = sheet.Slicers.Add(pivot, "F1", "Fruit");
        Slicer slicer = sheet.Slicers[slicerIndex];
        slicer.Caption = "Fruit Slicer";

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}