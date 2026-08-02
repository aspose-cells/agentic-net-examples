// Title: Refresh a PivotChart after modifying source data with Aspose.Cells (C#)
// Description: Load a workbook, update cells that feed a PivotTable, call RefreshPivotTables, loop through Worksheet.Charts to invoke RefreshPivotData on charts linked to a pivot source, and save the file. Ensures the PivotChart reflects the new data.
// Keywords: Aspose.Cells C# refresh PivotChart | RefreshPivotData example | Update PivotTable source programmatically | Pivot cache refresh Aspose.Cells | C# Excel PivotChart update | Aspose.Cells workbook modify cells | RefreshPivotTables worksheet
// Common Searches: how to refresh a PivotChart in Aspose.Cells C# | Aspose.Cells refresh pivot cache after data change | C# code to update PivotTable and PivotChart | RefreshPivotData Aspose.Cells example | programmatically refresh Excel PivotChart with Aspose
// Developer Intent: Update the workbook’s source values, refresh the associated PivotTable cache, refresh any PivotCharts, and save the updated file.
// Use Cases: Correct sales figures in an Excel report, then refresh all pivot visualizations before publishing. | Automate monthly KPI updates by changing key metrics, refreshing pivot structures, and generating a fresh workbook. | Batch‑process multiple Excel files to fix data errors, ensure pivot tables and charts are up‑to‑date, and overwrite the originals.
// AI Prompts: Generate C# code using Aspose.Cells that changes specific cell values, refreshes all PivotTables, refreshes only charts with a PivotSource, and saves the workbook. | Show how to iterate over Worksheet.Charts and call RefreshPivotData conditionally based on Chart.PivotSource. | Provide error‑handling patterns for RefreshPivotData when a chart is not linked to a pivot source.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

namespace PivotChartRefreshExample
{
    // Load a workbook, update cells that feed a PivotTable, call RefreshPivotTables, loop through Worksheet.Charts to invoke RefreshPivotData on charts linked to a pivot source, and save the file. Ensures the PivotChart reflects the new data.
    class Program
    {
        static void Main()
        {
            // Load the existing workbook that contains a PivotTable and a PivotChart
            Workbook workbook = new Workbook("input.xlsx");

            // Assume the first worksheet holds the data, pivot table and pivot chart
            Worksheet worksheet = workbook.Worksheets[0];

            // ----- Modify the source data that the pivot table uses -----
            // Example: change some values in the data range
            worksheet.Cells["B2"].PutValue(1500);
            worksheet.Cells["B3"].PutValue(2500);

            // ----- Refresh all PivotTables in the worksheet -----
            // This updates the pivot cache with the new source data
            worksheet.RefreshPivotTables();

            // ----- Refresh the PivotChart(s) that depend on the PivotTable -----
            // Iterate through all charts in the worksheet and refresh those that have a PivotSource
            foreach (Chart chart in worksheet.Charts)
            {
                if (!string.IsNullOrEmpty(chart.PivotSource))
                {
                    chart.RefreshPivotData();
                }
            }

            // Save the updated workbook
            workbook.Save("output.xlsx");
        }
    }
}
