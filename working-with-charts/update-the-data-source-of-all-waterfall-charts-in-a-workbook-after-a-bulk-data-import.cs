// Title: Refresh all Waterfall chart data sources after bulk import with Aspose.Cells for .NET
// Description: This example loads an existing workbook, changes the cells that feed the charts, walks through every worksheet, identifies Waterfall charts, and forces each chart to recalculate using ChartCalculateOptions.UpdateAllPoints so the visualisation reflects the new values before saving.
// Keywords: Aspose.Cells | C# | Waterfall chart | update chart data | ChartCalculateOptions | bulk data import | recalculate Excel chart | iterate worksheets | refresh charts programmatically | Excel automation
// Common Searches: How to recalculate Waterfall charts after cell updates with Aspose.Cells | Refresh Excel chart data programmatically in C# | ChartCalculateOptions.UpdateAllPoints usage | Iterate worksheets to refresh charts Aspose.Cells | Bulk import Excel data and update charts automatically
// Developer Intent: Programmatically recalculate every Waterfall chart in a workbook so it reflects newly imported data.
// Use Cases: After a nightly bulk load of sales numbers, automatically refresh all Waterfall charts for an up‑to‑date financial dashboard. | Generate monthly reports where the Waterfall visualisations must show the latest totals without manual editing. | In a batch processing pipeline that modifies a range of cells, ensure all Waterfall charts are refreshed before the workbook is distributed.
// AI Prompts: Write C# code using Aspose.Cells to refresh all Waterfall charts after modifying source cells. | Explain the effect of ChartCalculateOptions.UpdateAllPoints on Waterfall chart recalculation. | Show how to loop through worksheets and recalculate only Waterfall charts in an Excel file.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace UpdateWaterfallCharts
{
    // This example loads an existing workbook, changes the cells that feed the charts, walks through every worksheet, identifies Waterfall charts, and forces each chart to recalculate using ChartCalculateOptions.UpdateAllPoints so the visualisation reflects the new values before saving.
    class Program
    {
        static void Main()
        {
            // Load the workbook that contains Waterfall charts.
            // (Replace the file path with the actual location of your workbook.)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // -----------------------------------------------------------------
            // Assume a bulk data import has already been performed here.
            // For demonstration, we simply modify some cells that the charts use.
            // -----------------------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Cells["B2"].PutValue(1500);   // Example of updated data
            dataSheet.Cells["B3"].PutValue(2500);
            dataSheet.Cells["B4"].PutValue(1800);

            // Iterate through all worksheets and their charts.
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                for (int i = 0; i < sheet.Charts.Count; i++)
                {
                    Chart chart = sheet.Charts[i];

                    // Process only Waterfall charts.
                    if (chart.Type == ChartType.Waterfall)
                    {
                        // Recalculate the chart so that it picks up the latest cell values.
                        // UpdateAllPoints ensures that all data points are refreshed.
                        chart.Calculate(new ChartCalculateOptions { UpdateAllPoints = true });
                    }
                }
            }

            // Save the updated workbook.
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}
