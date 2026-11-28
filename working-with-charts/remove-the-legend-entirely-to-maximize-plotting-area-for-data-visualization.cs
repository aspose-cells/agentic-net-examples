// Title: Remove the legend and expand the plot area of a column chart using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# hide chart legend and enlarge plot area | how to turn off legend for a column chart with Aspose.Cells .NET | manually set plot area dimensions in Aspose.Cells chart C# | remove legend to maximize chart area Aspose.Cells example | configure ShowLegend property false in Aspose.Cells chart
// Tags: Aspose.Cells chart legend visibility | Aspose.Cells plot area manual sizing | Aspose.Cells column chart customization | Aspose.Cells ShowLegend property usage | Aspose.Cells workbook chart export to XLSX

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a column chart, disables the legend by setting ShowLegend to false, manually adjusts the plot area to occupy 95% of the chart dimensions, and saves the file as ChartWithoutLegend.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // -----------------------------------------------------------------
            // 1. Create a new workbook and get the first worksheet.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // -----------------------------------------------------------------
            // 2. Populate sample data for the chart.
            // -----------------------------------------------------------------
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // -----------------------------------------------------------------
            // 3. Add a column chart to the worksheet.
            // -----------------------------------------------------------------
            // Parameters: chart type, upper‑left row, upper‑left column,
            // lower‑right row, lower‑right column
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = worksheet.Charts[chartIndex];

            // Bind data series and categories.
            chart.NSeries.Add("B2:B4", true);          // Values
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // -----------------------------------------------------------------
            // 4. Remove the legend completely to maximize the plotting area.
            // -----------------------------------------------------------------
            // The Chart class exposes a boolean property ShowLegend.
            // Setting it to false hides the legend and frees the space.
            chart.ShowLegend = false;

            // Optional: make the plot area occupy most of the chart area.
            // Disable automatic sizing and set the width/height ratios.
            chart.PlotArea.IsAutomaticSize = false;
            chart.PlotArea.WidthRatioToChart = 0.95;   // 95 % of chart width
            chart.PlotArea.HeightRatioToChart = 0.95;  // 95 % of chart height

            // -----------------------------------------------------------------
            // 5. Save the workbook.
            // -----------------------------------------------------------------
            string outputPath = "ChartWithoutLegend.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Chart created without legend and saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Report any unexpected errors.
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
