// Title: C# – Apply a Value‑Greater‑Than Filter to a Pivot Table with Aspose.Cells
// Description: This example creates a workbook, adds product and sales data, builds a pivot table, and uses PivotFilterCollection.AddValueFilter with PivotFilterType.ValueGreaterThan to show only rows where the Sales amount exceeds a defined threshold. The pivot cache is refreshed, the data recalculated, and the file saved as an XLSX workbook.
// Keywords: Aspose.Cells pivot filter C# | ValueGreaterThan pivot table | filter pivot rows by sales | Aspose.Cells .NET example | apply numeric filter to pivot | C# Excel pivot value filter | Aspose.Cells PivotTable API
// Common Searches: Aspose.Cells filter pivot table values greater than | C# add value greater than filter to Aspose.Cells pivot | how to show only sales over a threshold in Aspose.Cells | PivotFilterCollection AddValueFilter example | Aspose.Cells pivot table numeric filter
// Developer Intent: Show only pivot table items whose Sales value is higher than a specified threshold.
// Use Cases: Generate a sales report that lists only products with revenue above $100. | Create a dashboard that automatically hides low‑performing items in a pivot view. | Export a filtered pivot table to Excel for stakeholder distribution.
// AI Prompts: How can I change the sales threshold at runtime and refresh the pivot table in Aspose.Cells? | Provide code to remove a ValueGreaterThan filter from an Aspose.Cells pivot table. | Explain how to combine a value filter with a label filter in a pivot table using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;

namespace AsposeCellsExamples
{
    // This example creates a workbook, adds product and sales data, builds a pivot table, and uses PivotFilterCollection.AddValueFilter with PivotFilterType.ValueGreaterThan to show only rows where the Sales amount exceeds a defined threshold. The pivot cache is refreshed, the data recalculated, and the file saved as an XLSX workbook.
    public class PivotValueFilterGreaterThanDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Product";
                sheet.Cells["B1"].Value = "Sales";

                sheet.Cells["A2"].Value = "WidgetA";
                sheet.Cells["B2"].Value = 120.0;

                sheet.Cells["A3"].Value = "WidgetB";
                sheet.Cells["B3"].Value = 85.0;

                sheet.Cells["A4"].Value = "WidgetC";
                sheet.Cells["B4"].Value = 200.0;

                sheet.Cells["A5"].Value = "WidgetD";
                sheet.Cells["B5"].Value = 45.0;

                // Create a pivot table based on the data range A1:B5, place it at D2
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D2", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add "Product" as a row field and "Sales" as a data field
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

                // Define the sales threshold
                double salesThreshold = 100.0;

                // Apply a value filter to show only items where Sales > salesThreshold
                PivotFilterCollection filters = pivotTable.PivotFilters;
                filters.AddValueFilter(
                    baseFieldIndex: 0,
                    valueFieldIndex: 1,
                    type: PivotFilterType.ValueGreaterThan,
                    value1: salesThreshold,
                    value2: 0.0 // Ignored for "greater than" filter
                );

                // Refresh the pivot cache and recalculate data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook
                workbook.Save("PivotValueFilterGreaterThanDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main()
        {
            Run();
        }
    }
}
