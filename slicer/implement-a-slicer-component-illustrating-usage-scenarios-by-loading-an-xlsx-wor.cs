using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

namespace SlicerDemoApp
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // Scenario 1: Add a slicer based on a ListObject (table)
            // -------------------------------------------------
            // Add and name the data worksheet
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["A4"].PutValue("A");
            dataSheet.Cells["B4"].PutValue(30);

            // Create a table (ListObject) over the data range
            int tableIdx = dataSheet.ListObjects.Add(0, 0, 3, 1, true);
            ListObject table = dataSheet.ListObjects[tableIdx];

            // Add a slicer for the first column of the table at cell D2
            int slicerIdxTable = dataSheet.Slicers.Add(table, 0, "D2");
            Slicer tableSlicer = dataSheet.Slicers[slicerIdxTable];
            tableSlicer.Caption = "Category Slicer";
            tableSlicer.StyleType = SlicerStyleType.SlicerStyleLight2;

            // -------------------------------------------------
            // Scenario 2: Add a slicer based on a PivotTable
            // -------------------------------------------------
            // Add a new worksheet for the pivot
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");

            // Populate pivot source data (reuse the same data)
            pivotSheet.Cells["A1"].PutValue("Category");
            pivotSheet.Cells["B1"].PutValue("Value");
            pivotSheet.Cells["A2"].PutValue("A");
            pivotSheet.Cells["B2"].PutValue(10);
            pivotSheet.Cells["A3"].PutValue("B");
            pivotSheet.Cells["B3"].PutValue(20);
            pivotSheet.Cells["A4"].PutValue("A");
            pivotSheet.Cells["B4"].PutValue(30);

            // Create a pivot table based on the data range
            int pivotIdx = pivotSheet.PivotTables.Add("=PivotSheet!A1:B4", "D5", "CategoryPivot");
            PivotTable pivot = pivotSheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer for the pivot field "Category" at cell G2
            int slicerIdxPivot = pivotSheet.Slicers.Add(pivot, "G2", "Category");
            Slicer pivotSlicer = pivotSheet.Slicers[slicerIdxPivot];
            pivotSlicer.Caption = "Pivot Category Slicer";
            pivotSlicer.StyleType = SlicerStyleType.SlicerStyleLight1;

            // Refresh slicers to ensure they reflect current data
            tableSlicer.Refresh();
            pivotSlicer.Refresh();

            // Save the modified workbook
            workbook.Save("SlicerDemo_Output.xlsx");
        }
    }
}