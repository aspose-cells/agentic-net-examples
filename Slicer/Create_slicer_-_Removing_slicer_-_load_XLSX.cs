using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is at least one pivot table; if not, create a simple one
        if (sheet.PivotTables.Count == 0)
        {
            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Fruit";
            sheet.Cells["B1"].Value = "Quantity";
            sheet.Cells["A2"].Value = "Apple";
            sheet.Cells["B2"].Value = 10;
            sheet.Cells["A3"].Value = "Orange";
            sheet.Cells["B3"].Value = 5;
            sheet.Cells["A4"].Value = "Banana";
            sheet.Cells["B4"].Value = 8;

            // Add a pivot table covering the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
        }

        // Retrieve the first (or only) pivot table
        PivotTable pivotTable = sheet.PivotTables[0];

        // Add a slicer for the "Fruit" field, placing its top‑left corner at cell E2
        int slicerIdx = sheet.Slicers.Add(pivotTable, "E2", "Fruit");

        // Access the slicer collection
        SlicerCollection slicers = sheet.Slicers;

        // Remove the slicer we just added using its index
        slicers.RemoveAt(slicerIdx);

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}