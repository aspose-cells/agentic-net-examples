using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class ApplySlicerStyle
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field
        pivot.CalculateData();

        // Add a slicer linked to the first base field of the pivot table
        int slicerIndex = sheet.Slicers.Add(pivot, "F1", 0);
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Apply the built‑in Light 1 slicer style
        slicer.StyleType = SlicerStyleType.SlicerStyleLight1;

        // Save the workbook with the styled slicer
        workbook.Save("SlicerStyleLight1.xlsx");
    }
}