// Title: Refresh a slicer’s cache after adding new source rows and verify the pivot table updates with Aspose.Cells for .NET (C#)
// AI Prompts: Insert a new product row into the source worksheet, call slicer.Refresh(), recalculate the pivot table, and print the updated row‑field item count. | Capture the pivot table row‑field item count before and after invoking Slicer.Refresh() to confirm that the cache refresh added the new item. | Save the workbook after the slicer refresh and output a boolean indicating whether the item count increased.
// Common Searches: Aspose.Cells C# refresh slicer after adding new data rows | how to update pivot table when slicer cache changes in Aspose.Cells for .NET | verify pivot table row items count after slicer.Refresh in C# | programmatically add source data and refresh linked slicer using Aspose.Cells
// Tags: slicer.Refresh Aspose.Cells C# | pivot table cache update after slicer refresh | add source rows programmatically Aspose.Cells | validate pivot table item count C# | Aspose.Cells slicer linked to pivot table

using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Demonstrates adding a new product row to the source worksheet, invoking Slicer.Refresh() to update the slicer cache, recalculating the linked pivot table, and confirming that the row‑field item count increased before saving the workbook.
class SlicerRefreshDemo
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Cells["A1"].PutValue("Product");
        dataSheet.Cells["B1"].PutValue("Sales");
        dataSheet.Cells["A2"].PutValue("Apple");
        dataSheet.Cells["B2"].PutValue(100);
        dataSheet.Cells["A3"].PutValue("Banana");
        dataSheet.Cells["B3"].PutValue(200);

        // Create a pivot table based on the data
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
        int pivotIndex = pivotSheet.PivotTables.Add("A1:B3", "C3", "PivotTable1");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales as data field
        pivotTable.CalculateData(); // Initial calculation

        // Add a slicer linked to the pivot table
        Worksheet slicerSheet = workbook.Worksheets.Add("Slicer");
        int slicerIndex = slicerSheet.Slicers.Add(pivotTable, "A1", "Product");
        Slicer slicer = slicerSheet.Slicers[slicerIndex];

        // Record the initial number of items in the pivot row field
        int initialItemCount = pivotTable.RowFields[0].PivotItems.Count;

        // Add a new item to the source data (which the slicer cache should pick up)
        dataSheet.Cells["A4"].PutValue("Orange");
        dataSheet.Cells["B4"].PutValue(150);

        // Refresh the slicer – this also refreshes the associated pivot table
        slicer.Refresh();

        // Ensure the pivot table recalculates after the refresh
        pivotTable.CalculateData();

        // Verify that the pivot table now contains the new item
        int refreshedItemCount = pivotTable.RowFields[0].PivotItems.Count;
        System.Console.WriteLine("Initial item count: " + initialItemCount);
        System.Console.WriteLine("Item count after refresh: " + refreshedItemCount);
        System.Console.WriteLine("Refresh successful: " + (refreshedItemCount > initialItemCount));

        // Save the workbook
        workbook.Save("SlicerRefreshResult.xlsx");
    }
}
