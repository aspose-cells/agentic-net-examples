// Title: Refresh an Aspose.Cells pivot table after modifying source data using C#
// AI Prompts: Generate C# code that changes cell values in a worksheet and then calls PivotTable.RefreshData() and PivotTable.CalculateData() to update the pivot results. | Show how to programmatically refresh the cache of an Aspose.Cells PivotTable after editing its source range in a .NET application. | Provide a step‑by‑step example that creates a workbook, adds a pivot table, updates source rows, refreshes the pivot, and saves the file.
// Common Searches: aspocells c# refresh pivot table after changing source cells | how to recalculate pivot cache programmatically with Aspose.Cells .NET | example of using RefreshData and CalculateData on Aspose.Cells pivot table | update worksheet data and refresh pivot table in Aspose.Cells C#
// Tags: Aspose.Cells refresh pivot cache C# | Aspose.Cells calculate pivot data .NET | update worksheet cells Aspose.Cells pivot | pivot table cache refresh after source change C# | programmatic pivot table recalculation Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook, adds a pivot table, modifies source cells, calls RefreshData and CalculateData to update the pivot, and saves the refreshed workbook.
class RefreshPivotDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate source data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["A4"].PutValue("Apple");
        sheet.Cells["B4"].PutValue(150);

        // Add a pivot table that uses the source range A1:B4 and places it at D1
        int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIdx];

        // Configure the pivot table: Product as row field, Sales as data field
        pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Column index 0 -> Product
        pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Column index 1 -> Sales

        // Initial calculation of the pivot table
        pivot.CalculateData();

        // ----- Modify the source data -----
        sheet.Cells["B2"].PutValue(120);   // Update Apple sales from 100 to 120
        sheet.Cells["A3"].PutValue("Apple"); // Change "Banana" to "Apple"

        // Refresh the pivot table's cache and recalculate the displayed data
        pivot.RefreshData();   // Gathers data from the updated source range
        pivot.CalculateData(); // Calculates the pivot results based on the refreshed cache

        // Save the workbook with the refreshed pivot table
        workbook.Save("RefreshedPivot.xlsx");
    }
}
