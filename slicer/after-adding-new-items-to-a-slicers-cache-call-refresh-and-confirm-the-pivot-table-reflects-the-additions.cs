// Title: Refresh an Aspose.Cells Slicer After Adding Source Rows and Verify Pivot Table Update (C#)
// Description: This C# example creates a workbook with product‑sales data, builds a pivot table, attaches a slicer, adds a new product row to the source sheet, calls slicer.Refresh() to update the slicer cache, recalculates the pivot, and confirms the new item appears by comparing the row‑field item count before and after the refresh. The workbook is then saved as an Excel file.
// Keywords: Aspose.Cells slicer refresh C# | update pivot table after adding data | slicer cache programmatically | pivot items count verification | dynamic Excel slicer Aspose | .NET Excel pivot refresh
// Common Searches: how to refresh a slicer in Aspose.Cells after inserting rows | C# code to sync slicer cache with pivot table | Aspose.Cells refresh slicer and recalculate pivot | verify new pivot items appear after slicer refresh .NET | programmatic slicer cache update Aspose.Cells
// Developer Intent: Programmatically refresh a slicer’s cache so the linked pivot table reflects newly added source rows.
// Use Cases: Add new rows to the data worksheet, invoke slicer.Refresh(), then call pivotTable.CalculateData() to keep the pivot in sync. | Compare the count of pivot row‑field items before and after data insertion to ensure the slicer cache refresh succeeded. | Generate Excel reports where slicers automatically stay up‑to‑date with changing source data by refreshing them on the fly.
// AI Prompts: Generate C# code using Aspose.Cells that adds rows to a source sheet, refreshes the associated slicer, and checks that the pivot table row items count increased. | Show how to handle slicer cache refresh and pivot recalculation when source data changes in Aspose.Cells for .NET. | Explain the steps to programmatically verify that a slicer’s cache has been updated after adding new items and how to trigger a pivot table refresh.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// This C# example creates a workbook with product‑sales data, builds a pivot table, attaches a slicer, adds a new product row to the source sheet, calls slicer.Refresh() to update the slicer cache, recalculates the pivot, and confirms the new item appears by comparing the row‑field item count before and after the refresh. The workbook is then saved as an Excel file.
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
        dataSheet.Cells["B3"].PutValue(150);

        // Create a worksheet for the pivot table
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

        // Record the initial number of pivot items (should be 2)
        int initialCount = pivotTable.RowFields[0].PivotItems.Count;
        Console.WriteLine("Initial Pivot Items Count: " + initialCount);

        // Add a new item to the source data (new product)
        dataSheet.Cells["A4"].PutValue("Orange");
        dataSheet.Cells["B4"].PutValue(200);

        // Refresh the slicer – this also refreshes the associated pivot table
        slicer.Refresh();

        // Ensure the pivot table recalculates after the refresh
        pivotTable.CalculateData();

        // Verify that the pivot now includes the new item
        int updatedCount = pivotTable.RowFields[0].PivotItems.Count;
        Console.WriteLine("Updated Pivot Items Count: " + updatedCount);

        // Save the workbook
        workbook.Save("SlicerRefreshResult.xlsx");
    }
}
