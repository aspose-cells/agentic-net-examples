// Title: C# – Refresh an Aspose.Cells PivotTable after Changing Source Data
// Description: Shows how to create a workbook, add a pivot table, alter source cells, and call PivotTable.RefreshData followed by CalculateData to refresh the cache and recalculate the pivot before saving the file.
// Keywords: Aspose.Cells | C# | RefreshData | PivotTable | update pivot cache | recalculate pivot | modify source data | Excel pivot programmatically | Aspose.Cells PivotTable Refresh | RefreshData vs RefreshAll
// Common Searches: Aspose.Cells refresh pivot table C# | How to update pivot cache after data change Aspose.Cells | PivotTable.RefreshData example | Refresh all pivot tables Aspose.Cells | Recalculate pivot after modifying source range C#
// Developer Intent: Update a PivotTable so it reflects recent changes to its source range.
// Use Cases: After adjusting sales figures in the source worksheet, invoke RefreshData and CalculateData to keep the summary report accurate. | When new rows or categories are added to the data sheet, refresh the pivot to ensure totals and groupings are up‑to‑date before distribution. | Automate a nightly job that modifies source data, refreshes every pivot in the workbook, and saves the updated file for downstream analytics.
// AI Prompts: Generate C# code using Aspose.Cells that modifies source cells, refreshes a specific PivotTable, and saves the workbook. | Explain the difference between PivotTable.RefreshData and Workbook.RefreshAll in Aspose.Cells with code examples for each scenario. | Create a reusable method that accepts a Workbook and pivot name, updates given source cells, refreshes the pivot cache, recalculates, and returns the updated workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDemo
{
    // Shows how to create a workbook, add a pivot table, alter source cells, and call PivotTable.RefreshData followed by CalculateData to refresh the cache and recalculate the pivot before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate source data for the pivot table
            dataSheet.Cells["A1"].PutValue("Product");
            dataSheet.Cells["B1"].PutValue("Sales");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["B2"].PutValue(120);
            dataSheet.Cells["A3"].PutValue("Banana");
            dataSheet.Cells["B3"].PutValue(80);
            dataSheet.Cells["A4"].PutValue("Apple");
            dataSheet.Cells["B4"].PutValue(150);
            dataSheet.Cells["A5"].PutValue("Banana");
            dataSheet.Cells["B5"].PutValue(200);

            // Add a new worksheet to host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

            // Create the pivot table using the source range A1:B5
            int pivotIndex = pivotSheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot fields (Product as row, Sales as data)
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Product
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Sales

            // Initial calculation to populate the pivot table
            pivotTable.CalculateData();

            // ----- Modify the source data -----
            dataSheet.Cells["B2"].PutValue(130); // Change Apple sales from 120 to 130
            dataSheet.Cells["A5"].PutValue("Apple"); // Change a Banana entry to Apple
            dataSheet.Cells["B5"].PutValue(210); // Adjust its sales

            // Refresh the pivot table to reflect the updated source data
            pivotTable.RefreshData();   // Gather data from the source into the pivot cache
            pivotTable.CalculateData(); // Recalculate the pivot table based on the refreshed cache

            // Save the workbook with the refreshed pivot table
            workbook.Save("PivotTable_Refreshed.xlsx");
        }
    }
}
