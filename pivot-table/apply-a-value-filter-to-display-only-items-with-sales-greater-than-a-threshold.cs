// Title: C# Aspose.Cells – Show Only Items with Sales > Threshold Using Pivot Table ValueGreaterThan Filter
// Description: Creates a workbook, inserts product‑sales data, builds a pivot table, and applies PivotFilterCollection.AddValueFilter with PivotFilterType.ValueGreaterThan so that only rows where Sales exceeds a defined threshold are displayed, then saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | ValueGreaterThan | PivotFilterCollection | sales threshold filter | filter pivot rows | Excel automation | Aspose.Cells example
// Common Searches: Aspose.Cells filter pivot table rows greater than value | C# add ValueGreaterThan filter to pivot table | Aspose.Cells PivotFilterType.ValueGreaterThan example | how to show only sales > 100 in Aspose.Cells pivot | filter pivot table by numeric threshold using Aspose.Cells
// Developer Intent: Display only those pivot‑table items whose Sales figure is higher than a specified numeric threshold.
// Use Cases: Produce a sales report that lists only products with revenue above a set limit. | Create a dynamic Excel dashboard that automatically hides low‑performing items. | Export a filtered pivot view for downstream analysis without low‑sales entries.
// AI Prompts: Write C# code with Aspose.Cells that adds a ValueGreaterThan filter to a pivot table based on a variable threshold. | Explain the parameters of PivotFilterCollection.AddValueFilter when using PivotFilterType.ValueGreaterThan. | Show how to change the sales threshold at runtime, refresh the pivot table, and save the updated workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Creates a workbook, inserts product‑sales data, builds a pivot table, and applies PivotFilterCollection.AddValueFilter with PivotFilterType.ValueGreaterThan so that only rows where Sales exceeds a defined threshold are displayed, then saves the workbook.
    public class PivotValueFilterGreaterThanDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            worksheet.Cells["A1"].Value = "Product";
            worksheet.Cells["B1"].Value = "Sales";

            worksheet.Cells["A2"].Value = "WidgetA";
            worksheet.Cells["B2"].Value = 150;
            worksheet.Cells["A3"].Value = "WidgetB";
            worksheet.Cells["B3"].Value = 80;
            worksheet.Cells["A4"].Value = "WidgetC";
            worksheet.Cells["B4"].Value = 220;
            worksheet.Cells["A5"].Value = "WidgetD";
            worksheet.Cells["B5"].Value = 45;

            // Add a pivot table based on the data range A1:B5, place it at D3
            int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

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
                value2: 0); // value2 is ignored for ValueGreaterThan

            // Refresh and calculate the pivot table to apply the filter
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            string outputPath = "PivotValueFilterGreaterThanDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
