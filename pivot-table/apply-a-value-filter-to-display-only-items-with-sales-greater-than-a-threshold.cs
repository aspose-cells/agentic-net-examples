// Title: Apply a ValueGreaterThan filter to a pivot table in Aspose.Cells for .NET (C#) to display only products with sales above a threshold
// AI Prompts: Generate an Excel workbook with a pivot table that shows only rows where the Sales column exceeds a given value using Aspose.Cells in C#. | Add a PivotFilterCollection.ValueGreaterThan filter to the row field of a pivot table, then refresh and calculate the pivot cache with Aspose.Cells for .NET. | Write a C# program that creates a pivot table from a data range, applies a sales‑greater‑than filter, and saves the filtered result as an .xlsx file.
// Common Searches: Aspose.Cells C# apply value greater than filter on pivot table sales column | filter pivot table rows by sales amount using Aspose.Cells for .NET | example code for PivotFilterType.ValueGreaterThan in C# Aspose.Cells | show only products with sales over 100 in an Aspose.Cells pivot table | refresh pivot cache after applying value filter Aspose.Cells C#
// Tags: Aspose.Cells pivot table value filter C# | PivotFilterType.ValueGreaterThan Aspose.Cells | filter pivot rows by sales threshold C# | refresh pivot cache Aspose.Cells | save filtered pivot table to Excel C#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // The example creates a workbook with product and sales data, builds a pivot table, applies a PivotFilterType.ValueGreaterThan filter on the Sales field to keep only items with sales above a defined threshold, refreshes and calculates the pivot, and saves the filtered pivot table to an Excel file.
    public class ApplyValueFilterDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data: Product names and their sales figures
                worksheet.Cells["A1"].Value = "Product";
                worksheet.Cells["B1"].Value = "Sales";
                worksheet.Cells["A2"].Value = "WidgetA";
                worksheet.Cells["B2"].Value = 120.0;
                worksheet.Cells["A3"].Value = "WidgetB";
                worksheet.Cells["B3"].Value = 85.0;
                worksheet.Cells["A4"].Value = "WidgetC";
                worksheet.Cells["B4"].Value = 200.0;
                worksheet.Cells["A5"].Value = "WidgetD";
                worksheet.Cells["B5"].Value = 45.0;

                // Define the data range for the pivot table (including headers)
                string dataRange = "A1:B5";

                // Add a pivot table to the worksheet
                int pivotIndex = worksheet.PivotTables.Add(dataRange, "D3", "SalesPivot");
                PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

                // Add the Product field to the Row area
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0); // 0 = column A (Product)

                // Add the Sales field to the Data area
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // 1 = column B (Sales)

                // Define the sales threshold
                double salesThreshold = 100.0;

                // Apply a value filter to show only items with Sales > salesThreshold
                PivotFilterCollection filters = pivotTable.PivotFilters;
                filters.AddValueFilter(
                    baseFieldIndex: 0,
                    valueFieldIndex: 1,
                    type: PivotFilterType.ValueGreaterThan,
                    value1: salesThreshold,
                    value2: 0);

                // Refresh the pivot cache and calculate data to apply the filter
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook with the filtered pivot table
                workbook.Save("ApplyValueFilterDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ApplyValueFilterDemo.Run();
        }
    }
}
