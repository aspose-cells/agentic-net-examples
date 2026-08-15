// Title: Aspose.Cells .NET: Create a Pivot Table with Custom Numeric Grouping for Sales Amounts
// Description: Demonstrates how to build a workbook, add sample sales data, generate a pivot table, place the Sales field in Row and Data areas, and apply a numeric range grouping (0‑10,000 with a 2,000 interval) using PivotField.GroupBy without creating a new field. The pivot is refreshed, recalculated, and saved as an Excel file.
// Keywords: Aspose.Cells pivot table numeric grouping | C# custom range grouping | PivotField.GroupBy example | group sales amounts by interval .NET | Aspose.Cells custom numeric range | Excel pivot numeric range grouping | Aspose.Cells .NET tutorial
// Common Searches: Aspose.Cells how to group numeric field in pivot table | C# pivot table custom range grouping Aspose | Set start end interval for pivot numeric grouping .NET | Apply numeric range grouping without new field Aspose.Cells | Pivot table sales grouping 2000 interval C#
// Developer Intent: The developer needs to generate a pivot table and apply a custom numeric range grouping to the Sales column programmatically.
// Use Cases: Summarize sales totals in fixed intervals (0‑1999, 2000‑3999, etc.) for financial dashboards. | Produce a compact report that shows aggregated sales per range without adding extra worksheet columns. | Programmatically adjust grouping intervals to explore different sales distribution views from the same data set.
// AI Prompts: Show how to change the grouping interval to 5,000 and update the pivot table using Aspose.Cells. | Provide code to list each generated numeric group label after calling GroupBy. | Explain how to create a separate pivot field for the grouped sales values instead of grouping the original field.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace CustomNumericGroupingDemo
{
    // Demonstrates how to build a workbook, add sample sales data, generate a pivot table, place the Sales field in Row and Data areas, and apply a numeric range grouping (0‑10,000 with a 2,000 interval) using PivotField.GroupBy without creating a new field. The pivot is refreshed, recalculated, and saved as an Excel file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and a worksheet for source data
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate sample sales data
                dataSheet.Cells["A1"].PutValue("Sales");
                double[] salesValues = { 500, 1500, 2500, 3500, 4500, 5500, 6500, 7500, 8500, 9500 };
                for (int i = 0; i < salesValues.Length; i++)
                {
                    dataSheet.Cells[i + 1, 0].PutValue(salesValues[i]);
                }

                // Add a worksheet to host the pivot table
                Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

                // Create a pivot table based on the sales data range
                int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:A11", "A1", "SalesPivot");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Add the Sales field to the Row area
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Sales");

                // Add the Sales field to the Data area to show sum of each group
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh the pivot table so that fields are materialized
                pivotTable.RefreshData();

                // Retrieve the pivot field that represents the Sales column
                PivotField salesField = pivotTable.RowFields[0];

                // Define custom numeric grouping: 0‑9999 with 2000 interval
                double start = 0.0;
                double end = 10000.0;
                double interval = 2000.0;

                // Apply the numeric range grouping; false indicates we do NOT create a new field
                salesField.GroupBy(start, end, interval, false);

                // Verify grouping settings (guard against null)
                if (salesField.GroupSettings is PivotNumbericRangeGroupSettings groupSettings)
                {
                    Console.WriteLine("Numeric grouping applied. Interval = " + groupSettings.Interval);
                }
                else
                {
                    Console.WriteLine("Grouping settings were not applied.");
                }

                // Recalculate the pivot table to reflect grouping
                pivotTable.CalculateData();

                // Save the workbook with the pivot table and custom grouping
                workbook.Save("CustomNumericGrouping.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
