// Title: Clear a Pivot Table Row Field Filter using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to remove a label filter from a pivot table row field with PivotField.ClearFilter(), refresh the cache, recalculate data, and save the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells ClearFilter | C# pivot table remove row filter | Aspose.Cells PivotField.ClearFilter example | reset pivot row field filter .NET | Aspose.Cells PivotTable filter programmatically | remove label filter Aspose.Cells
// Common Searches: Aspose.Cells how to clear pivot row filter C# | remove label filter from pivot table using Aspose.Cells | PivotField.ClearFilter usage in .NET | reset pivot table row field programmatically | Aspose.Cells clear filter example
// Developer Intent: Programmatically clear an applied filter on a specific pivot table row field so that all row items become visible again.
// Use Cases: After filtering a pivot report to a single region, clear the filter to show the full list of regions without rebuilding the pivot. | Implement a toggle button that applies a label filter and later removes it to let users switch between focused and full views. | Refresh pivot calculations after clearing filters to ensure totals and subtotals reflect the complete dataset.
// AI Prompts: Write C# code that uses Aspose.Cells to clear a label filter on a pivot table row field and then refreshes the pivot data. | Explain the role of PivotField.ClearFilter in Aspose.Cells and the steps required to recalculate a pivot table after clearing a filter. | Show how to iterate over all row fields in a pivot table and clear their filters using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsExamples
{
    // Demonstrates how to remove a label filter from a pivot table row field with PivotField.ClearFilter(), refresh the cache, recalculate data, and save the workbook using Aspose.Cells for .NET.
    public class ClearRowFieldFilterDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Region";
                sheet.Cells["B1"].Value = "Sales";
                sheet.Cells["A2"].Value = "North";
                sheet.Cells["B2"].Value = 1200;
                sheet.Cells["A3"].Value = "South";
                sheet.Cells["B3"].Value = 850;
                sheet.Cells["A4"].Value = "East";
                sheet.Cells["B4"].Value = 950;
                sheet.Cells["A5"].Value = "West";
                sheet.Cells["B5"].Value = 1100;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D2", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add the "Region" column as a row field
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);

                // Add the "Sales" column as a data field
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

                // Apply a filter on the row field (Region) to show only "North"
                PivotField rowField = pivotTable.RowFields[0];
                rowField.FilterByLabel(PivotFilterType.CaptionEqual, "North", null);
                // Refresh pivot cache and recalculate data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Clear the filter on the row field, restoring all regions
                rowField.ClearFilter();
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Save the workbook to verify the filter has been cleared
                workbook.Save("ClearRowFieldFilterDemo.xlsx");
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
