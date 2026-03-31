using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class SlicerManagementDemo
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("Input.xlsx");

        // Get the first worksheet
        Worksheet ws = workbook.Worksheets[0];

        // -------------------------------------------------
        // Prepare sample data and create a table (ListObject)
        // -------------------------------------------------
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["A2"].PutValue("A");
        ws.Cells["B2"].PutValue(10);
        ws.Cells["A3"].PutValue("B");
        ws.Cells["B3"].PutValue(20);
        ws.Cells["A4"].PutValue("A");
        ws.Cells["B4"].PutValue(30);
        ws.Cells["A5"].PutValue("B");
        ws.Cells["B5"].PutValue(40);

        // Add a ListObject that covers the data range A1:B5
        int tableIdx = ws.ListObjects.Add(0, 0, 4, 1, true);
        ListObject table = ws.ListObjects[tableIdx];

        // -------------------------------------------------
        // Add slicers that are based on the table
        // -------------------------------------------------
        // 1. Using column index and destination cell name
        int slicerIdx1 = ws.Slicers.Add(table, 0, "D2");

        // 2. Using ListColumn object and destination cell name
        int slicerIdx2 = ws.Slicers.Add(table, table.ListColumns[0], "G2");

        // 3. Using column index and destination cell name (cell F3)
        int slicerIdx3 = ws.Slicers.Add(table, 0, "F3");

        // Access the first slicer and customize its appearance
        Slicer slicer1 = ws.Slicers[slicerIdx1];
        slicer1.Caption = "Category Slicer";
        slicer1.StyleType = SlicerStyleType.SlicerStyleLight2;
        slicer1.LockedPosition = true; // Prevent user from moving the slicer
        slicer1.Refresh(); // Ensure the slicer reflects current data

        // -------------------------------------------------
        // Create a pivot table based on the same data range
        // -------------------------------------------------
        PivotTableCollection pivots = ws.PivotTables;
        int pivotIdx = pivots.Add("=Sheet1!A1:B5", "J2", "Pivot1");
        PivotTable pivot = pivots[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Category");
        pivot.AddFieldToArea(PivotFieldType.Data, "Value");
        pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
        pivot.RefreshData();
        pivot.CalculateData();

        // -------------------------------------------------
        // Add slicers that are based on the pivot table
        // -------------------------------------------------
        // 1. Using destination cell name and base field name
        int slicerIdx4 = ws.Slicers.Add(pivot, "L2", "Category");

        // 2. Using row/column indices and base field name (row=5, column=12 => cell M6)
        int slicerIdx5 = ws.Slicers.Add(pivot, 5, 12, "Category");

        // Customize the pivot slicer
        Slicer slicerPivot = ws.Slicers[slicerIdx4];
        slicerPivot.Caption = "Pivot Category";
        slicerPivot.StyleType = SlicerStyleType.SlicerStyleLight3;

        // -------------------------------------------------
        // Demonstrate slicer removal and clearing
        // -------------------------------------------------
        // Remove the second table slicer (index slicerIdx2)
        ws.Slicers.RemoveAt(slicerIdx2);

        // Clear all remaining slicers from the worksheet
        ws.Slicers.Clear();

        // -------------------------------------------------
        // Save the workbook with the modifications
        // -------------------------------------------------
        workbook.Save("Output.xlsx");
    }
}