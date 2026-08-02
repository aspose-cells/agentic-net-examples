using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerRefreshDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet (data sheet)
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate initial source data for the pivot table
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Sales");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(100);
            dataSheet.Cells["A3"].PutValue("Banana");
            dataSheet.Cells["B3"].PutValue(200);

            // Add a worksheet to host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotSheet");
            int pivotIndex = pivotSheet.PivotTables.Add("A1:B3", "C3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot table (row field = Product, data field = Sales)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);
            pivotTable.CalculateData();

            // Add a slicer linked to the pivot table for the "Product" field
            Worksheet slicerSheet = workbook.Worksheets.Add("SlicerSheet");
            int slicerIndex = slicerSheet.Slicers.Add(pivotTable, "A1", "Product");
            Slicer slicer = slicerSheet.Slicers[slicerIndex];

            // Record the initial number of items in the pivot row field (should be 2)
            int initialCount = pivotTable.RowFields[0].PivotItems.Count;

            // Add a new item to the source data (new product "Orange")
            dataSheet.Cells["A4"].PutValue("Orange");
            dataSheet.Cells["B4"].PutValue(150);

            // Refresh the slicer; this also refreshes the underlying pivot table
            slicer.Refresh();

            // Ensure the pivot table recalculates after the refresh
            pivotTable.CalculateData();

            // Verify that the pivot table now includes the new item
            int updatedCount = pivotTable.RowFields[0].PivotItems.Count;
            Console.WriteLine("Pivot row items increased: " + (updatedCount > initialCount));

            // Save the workbook (lifecycle rule)
            workbook.Save("SlicerRefreshResult.xlsx");
        }
    }
}