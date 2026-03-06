using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load an existing XLSX file
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one pivot table; create a simple one if none exist
        if (sheet.PivotTables.Count == 0)
        {
            // Sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(30);

            // Add a pivot table covering the sample data
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        }

        // Retrieve the first pivot table in the worksheet
        PivotTable pivotTable = sheet.PivotTables[0];

        // Add a slicer linked to the first base field of the pivot table
        // Destination cell for the slicer is "F1"
        int slicerIdx = sheet.Slicers.Add(pivotTable, "F1", pivotTable.BaseFields[0].Name);
        Slicer slicer = sheet.Slicers[slicerIdx];

        // Optional: customize slicer appearance
        slicer.Caption = "Category Slicer";
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}