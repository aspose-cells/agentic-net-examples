using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure a pivot table exists; create one if necessary
        PivotTable pivot;
        if (sheet.PivotTables.Count == 0)
        {
            int pivotIdx = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
            pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);
            pivot.AddFieldToArea(PivotFieldType.Data, 1);
        }
        else
        {
            pivot = sheet.PivotTables[0];
        }

        // Add a slicer linked to the first base field of the pivot table
        int slicerIdx = sheet.Slicers.Add(pivot, "F1", 0);
        Slicer slicer = sheet.Slicers[slicerIdx];

        // Optional: customize slicer appearance
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
        slicer.Caption = "Sample Slicer";

        // Save the workbook with the new slicer
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}