// Title: Log PivotTable Row Count After Refresh with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a PivotTable, refresh it, and read the count of distinct row items using the RowFields collection. The sample logs the row count before and after modifying the source data and saves the workbook, enabling performance monitoring and diagnostics.
// Keywords: Aspose.Cells | C# | PivotTable row count | RefreshData | PivotItems count | Excel automation | performance diagnostics | log pivot rows | data monitoring | Excel pivot refresh
// Common Searches: Aspose.Cells get pivot table row count | C# count rows in PivotTable after refresh | How to log pivot row items with Aspose | Retrieve distinct row items from Aspose.Cells pivot | Monitor pivot table size for performance
// Developer Intent: Learn how to obtain and log the number of distinct row items in an Aspose.Cells PivotTable after each RefreshData call for monitoring data changes and performance.
// Use Cases: Record row count before and after source updates to detect unexpected growth. | Send pivot row metrics to a monitoring system or log file for performance analysis. | Validate that newly added categories appear in the PivotTable by comparing row counts across refresh cycles.
// AI Prompts: Generate C# code using Aspose.Cells that logs the row count of a PivotTable after every RefreshData call and writes the values to a text file. | Show how to capture row counts for multiple row fields in a PivotTable and output the results as JSON for downstream processing. | Explain how to integrate PivotTable row‑count monitoring into an existing .NET performance logging framework with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableRowCountLogging
{
    // Demonstrates how to create a workbook, add a PivotTable, refresh it, and read the count of distinct row items using the RowFields collection. The sample logs the row count before and after modifying the source data and saves the workbook, enabling performance monitoring and diagnostics.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Clothing");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Food");
            sheet.Cells["B4"].PutValue(150);
            sheet.Cells["A5"].PutValue("Electronics");
            sheet.Cells["B5"].PutValue(200);

            // Add a pivot table based on the source data
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Category as row field, Amount as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Refresh the pivot table to ensure it reflects the current source data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Log the number of rows generated (i.e., distinct row items)
            int rowCount = pivotTable.RowFields[0].PivotItems.Count;
            Console.WriteLine($"Row count after first refresh: {rowCount}");

            // Modify source data to add a new category
            sheet.Cells["A6"].PutValue("Furniture");
            sheet.Cells["B6"].PutValue(300);

            // Refresh again to capture the new data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Log the updated row count
            int updatedRowCount = pivotTable.RowFields[0].PivotItems.Count;
            Console.WriteLine($"Row count after second refresh: {updatedRowCount}");

            // Save the workbook (lifecycle rule)
            workbook.Save("PivotTableRowCountDemo.xlsx");
        }
    }
}
