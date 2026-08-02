using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class SlicerVersionedSaveDemo
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Cells["A1"].PutValue("Product");
        dataSheet.Cells["B1"].PutValue("Sales");
        dataSheet.Cells["A2"].PutValue("Apple");
        dataSheet.Cells["B2"].PutValue(120);
        dataSheet.Cells["A3"].PutValue("Banana");
        dataSheet.Cells["B3"].PutValue(150);
        dataSheet.Cells["A4"].PutValue("Orange");
        dataSheet.Cells["B4"].PutValue(90);

        // Create a pivot table based on the data
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
        int pivotIndex = pivotSheet.PivotTables.Add("A1:B4", "C3", "PivotTable1");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column

        // Add a slicer linked to the pivot table
        Worksheet slicerSheet = workbook.Worksheets.Add("Slicer");
        int slicerIndex = slicerSheet.Slicers.Add(pivotTable, "A1", "Product");
        Slicer slicer = slicerSheet.Slicers[slicerIndex];

        // Save version 0 (initial state)
        workbook.Save("Workbook_Version_0.xlsx");

        // ---- Modification 1: Change slicer caption ----
        slicer.Caption = "Product Filter";
        slicer.Refresh(); // Refresh slicer and associated pivot table
        workbook.Save("Workbook_Version_1.xlsx");

        // ---- Modification 2: Lock slicer position ----
        slicer.LockedPosition = true;
        slicer.Refresh();
        workbook.Save("Workbook_Version_2.xlsx");

        // ---- Modification 3: Change number of columns in slicer ----
        slicer.NumberOfColumns = 3;
        slicer.Refresh();
        workbook.Save("Workbook_Version_3.xlsx");

        // Clean up resources
        workbook.Dispose();
    }
}