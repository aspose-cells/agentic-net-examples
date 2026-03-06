using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class ManageSlicerSample
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("Input.xlsx");

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is a pivot table; if not, create a simple one for demonstration
        if (sheet.PivotTables.Count == 0)
        {
            // Populate sample data
            sheet.Cells["A1"].PutValue("Fruit");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a pivot table using the Add(PivotTable, string, string) overload
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
        }

        // Retrieve the first pivot table
        PivotTable pivotTable = sheet.PivotTables[0];

        // Add a slicer linked to the pivot table for the "Fruit" field,
        // placing the slicer's upper‑left corner at cell E1
        int slicerIdx = sheet.Slicers.Add(pivotTable, "E1", "Fruit");
        Slicer slicer = sheet.Slicers[slicerIdx];

        // Configure slicer appearance and behavior
        slicer.Caption = "Fruit Slicer";
        slicer.NumberOfColumns = 2;
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
        slicer.LockedPosition = false;

        // Refresh the slicer to ensure it reflects the current pivot data
        slicer.Refresh();

        // Save the modified workbook
        workbook.Save("Output.xlsx");
    }
}