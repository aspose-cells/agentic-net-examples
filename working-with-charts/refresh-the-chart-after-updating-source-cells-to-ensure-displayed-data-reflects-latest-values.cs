// Title: Refresh Aspose.Cells chart after modifying source cells (C# .NET)
// Description: Demonstrates how to update cell values, invoke chart.Calculate (or RefreshPivotData for PivotCharts) to sync the visual representation, and save the workbook with the refreshed chart.
// Keywords: Aspose.Cells chart refresh | chart.Calculate C# | RefreshPivotData Aspose.Cells | update chart data programmatically | .NET workbook chart cache | Aspose.Cells column chart example
// Common Searches: Aspose.Cells refresh chart after cell change | C# chart.Calculate vs RefreshPivotData | How to update Aspose.Cells chart data | Refresh chart cache in Aspose.Cells .NET
// Developer Intent: Synchronize a chart with the latest values written to its source range.
// Use Cases: Change a series of numeric cells and call chart.Calculate to display the new values before saving. | Reload a PivotChart after altering its underlying pivot table using chart.RefreshPivotData. | Perform bulk data edits across worksheets and refresh all linked charts in one pass.
// AI Prompts: Generate C# code that updates cells and refreshes an Aspose.Cells chart using chart.Calculate. | Explain when to prefer chart.Calculate over chart.RefreshPivotData in Aspose.Cells. | Show how to loop through multiple charts in a workbook and refresh each after bulk data updates.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartRefreshDemo
{
    // Demonstrates how to update cell values, invoke chart.Calculate (or RefreshPivotData for PivotCharts) to sync the visual representation, and save the workbook with the refreshed chart.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart and set its data range
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // -----------------------------------------------------------------
            // Update the source cells – the chart should reflect these changes.
            // -----------------------------------------------------------------
            sheet.Cells["B2"].PutValue(15);   // Change first value
            sheet.Cells["B3"].PutValue(25);   // Change second value
            sheet.Cells["B4"].PutValue(35);   // Change third value

            // Refresh the chart so that it uses the latest cell values.
            // For a regular chart, calling Calculate updates its internal cache.
            chart.Calculate();

            // Alternatively, if the chart is a PivotChart, you would use:
            // chart.RefreshPivotData();

            // Save the workbook (the chart now displays the updated data)
            workbook.Save("ChartRefreshed.xlsx", SaveFormat.Xlsx);
        }
    }
}
