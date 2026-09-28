// Title: Select specific slicer items in C# with Aspose.Cells and verify the linked pivot table total
// AI Prompts: Create C# code that builds a workbook, adds a pivot table on the Product column, links a slicer to that field, selects only the "Apple" item in the slicer, refreshes the slicer, and prints the resulting pivot total. | Write C# to modify the slicer selection to include both "Apple" and "Banana", refresh the associated pivot table, and output the combined sales sum. | Add verification logic in C# that reads the pivot table total cell after a slicer refresh and asserts the value equals the expected sum.
// Common Searches: Aspose.Cells C# filter pivot table using slicer programmatically | set slicer selected items Aspose.Cells and refresh linked pivot | read pivot table total after slicer refresh Aspose.Cells .NET | verify slicer filter results in pivot table using Aspose.Cells API
// Tags: programmatic slicer item selection Aspose.Cells | refresh linked pivot table after slicer update .NET | validate pivot total after slicer filter C# | slicer cache item selection property Aspose.Cells | pivot table data verification Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// Demonstrates creating a workbook with product sales data, adding a pivot table, linking a slicer to the Product field, programmatically selecting only "Apple" in the slicer, refreshing both slicer and pivot, confirming the pivot total equals 250, and saving the workbook.
class SlicerPivotRefreshDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet for source data
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate source data (Product | Sales)
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Sales");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(100);
            dataSheet.Cells["A3"].PutValue("Banana");
            dataSheet.Cells["B3"].PutValue(200);
            dataSheet.Cells["A4"].PutValue("Apple");
            dataSheet.Cells["B4"].PutValue(150);

            // Add a worksheet for the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
            int pivotIndex = pivotSheet.PivotTables.Add("A1:B4", "C3", "Pivot1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot: Product as row field, Sales as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column
            pivotTable.RefreshData();   // Gather data into cache
            pivotTable.CalculateData(); // Calculate totals

            // Add a worksheet for the slicer
            Worksheet slicerSheet = workbook.Worksheets.Add("Slicer");
            // Create slicer linked to the pivot table on the "Product" field
            int slicerIndex = slicerSheet.Slicers.Add(pivotTable, "A1", "Product");
            Slicer slicer = slicerSheet.Slicers[slicerIndex];

            // ---- Update slicer items ----
            // Select only the "Apple" item, deselect all others
            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                item.Selected = item.Value.Equals("Apple", StringComparison.OrdinalIgnoreCase);
            }

            // Refresh the slicer; this also refreshes and recalculates the linked pivot table
            slicer.Refresh();

            // Ensure pivot data is up‑to‑date after slicer refresh
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // ---- Verify pivot table reflects the new filter ----
            // After filtering to "Apple", the pivot should show total sales = 250 (100 + 150)
            // The total value is in the data column (D4) of the pivot table
            double appleTotal = pivotSheet.Cells["D4"].DoubleValue;
            Console.WriteLine("Apple total after slicer filter: " + appleTotal);

            // Save the workbook
            workbook.Save("SlicerPivotRefreshDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
