// Title: Create an Excel PivotTable with a custom tooltip using Aspose.Cells in C#
// AI Prompts: Write C# code that uses Aspose.Cells to generate a pivot table and assign a custom tooltip via AltTextDescription and AltTextTitle. | Show how to add alt text (tooltip) to an existing Aspose.Cells pivot table in a .NET workbook. | Provide a step‑by‑step example for setting a custom tooltip on a pivot table and refreshing its cache with Aspose.Cells for C#. | Demonstrate modifying a workbook to include a pivot table with a descriptive tooltip using Aspose.Cells APIs.
// Common Searches: how to set a custom tooltip for a pivot table with Aspose.Cells C# | Aspose.Cells C# pivot table AltTextDescription example | add alt text title to Excel pivot table using Aspose.Cells .NET | refresh pivot cache after changing tooltip Aspose.Cells | C# code sample for creating pivot table with tooltip in Aspose.Cells
// Tags: Aspose.Cells set pivot table tooltip | AltTextDescription Aspose.Cells C# | pivot table alt text title Aspose.Cells | refresh pivot cache Aspose.Cells | C# create pivot table with custom tooltip

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates creating a workbook, populating data, adding a pivot table, assigning custom tooltip text via AltTextDescription and AltTextTitle, refreshing the pivot cache, and saving the file as PivotTableWithCustomTooltip.xlsx using Aspose.Cells for .NET.
    public class PivotTableCustomTooltipDemo
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
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Region");
            worksheet.Cells["C1"].PutValue("Sales");

            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue("North");
            worksheet.Cells["C2"].PutValue(1200);

            worksheet.Cells["A3"].PutValue("Apple");
            worksheet.Cells["B3"].PutValue("South");
            worksheet.Cells["C3"].PutValue(800);

            worksheet.Cells["A4"].PutValue("Banana");
            worksheet.Cells["B4"].PutValue("North");
            worksheet.Cells["C4"].PutValue(1500);

            worksheet.Cells["A5"].PutValue("Banana");
            worksheet.Cells["B5"].PutValue("South");
            worksheet.Cells["C5"].PutValue(1100);

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = worksheet.PivotTables;
            int pivotIndex = pivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Add fields to the pivot table
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Set a custom tooltip (alt text description) for the pivot table
            pivotTable.AltTextDescription = "Shows sales distribution by product and region";
            pivotTable.AltTextTitle = "Sales Pivot Table";

            // Refresh the pivot cache and calculate the pivot table data
            pivotTable.RefreshData();      // Correct API to refresh cache
            pivotTable.CalculateData();

            // Save the workbook with the configured pivot table and tooltip
            workbook.Save("PivotTableWithCustomTooltip.xlsx");
        }
    }
}
