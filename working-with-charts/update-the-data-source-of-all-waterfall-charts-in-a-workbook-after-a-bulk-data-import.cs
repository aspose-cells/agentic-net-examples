// Title: Refresh Waterfall Charts After Bulk Data Import with Aspose.Cells for .NET
// Description: Loads a workbook, scans every worksheet, identifies Waterfall charts, and forces a full data refresh using Chart.Calculate with UpdateAllPoints=true. The updated workbook is then saved, ensuring charts reflect the newly imported data.
// Keywords: Aspose.Cells Waterfall chart refresh | update chart data source .NET | Chart.Calculate UpdateAllPoints | bulk data import Excel chart update | recalculate Excel charts programmatically
// Common Searches: how to refresh waterfall charts after data import asp.net | aspocells recalculate all charts in workbook | update Excel chart data source with Aspose.Cells C# | refresh Excel charts after bulk load using Aspose
// Developer Intent: Programmatically refresh every Waterfall chart in a workbook so it displays the latest imported values.
// Use Cases: Iterate through all worksheets and charts, detect Waterfall type, and call chart.Calculate with UpdateAllPoints=true to sync visuals with new data. | Save the workbook after the refresh to produce an output file containing updated Waterfall chart graphics.
// AI Prompts: Write C# code that uses Aspose.Cells to refresh all Waterfall charts in an Excel file after modifying the source range. | Explain the purpose of ChartCalculateOptions.UpdateAllPoints when recalculating charts in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace WaterfallChartRefreshDemo
{
    // Loads a workbook, scans every worksheet, identifies Waterfall charts, and forces a full data refresh using Chart.Calculate with UpdateAllPoints=true. The updated workbook is then saved, ensuring charts reflect the newly imported data.
    class Program
    {
        static void Main()
        {
            // Load the workbook that contains the Waterfall charts.
            // (Replace the file name with the actual path of your workbook.)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Assume that a bulk data import has already been performed.
            // If you need to import data programmatically, do it here before refreshing the charts.

            // Iterate through all worksheets in the workbook.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts on the current worksheet.
                for (int i = 0; i < sheet.Charts.Count; i++)
                {
                    Chart chart = sheet.Charts[i];

                    // Check if the chart is a Waterfall chart.
                    if (chart.Type == ChartType.Waterfall)
                    {
                        // Recalculate the chart data. Setting UpdateAllPoints to true ensures
                        // that all data points are refreshed according to the latest source data.
                        chart.Calculate(new ChartCalculateOptions { UpdateAllPoints = true });
                    }
                }
            }

            // Save the updated workbook.
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}
