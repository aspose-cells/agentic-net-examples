// Title: Create a column chart and change the legend from automatic positioning to a fixed top‑left location using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set chart legend to automatic position then move it to a specific location | how to place a chart legend at the top left corner using XRatioToChart in Aspose.Cells | C# Aspose.Cells column chart with custom legend coordinates | disable automatic legend sizing and set fixed X/Y ratios in Aspose.Cells chart | Aspose.Cells example for overriding legend position after automatic placement
// Tags: chart legend automatic positioning Aspose.Cells | set legend XRatioToChart Aspose.Cells | column chart generation Aspose.Cells .NET | override legend top left location Aspose.Cells | disable legend automatic size Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLegendDemo
{
    // Demonstrates creating a column chart, using the default automatic legend placement, then disabling automatic sizing and setting XRatioToChart and YRatioToChart to 0 to position the legend at the chart’s top‑left corner, and saving the workbook as ChartWithCustomLegend.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data for the chart
                cells["A1"].PutValue("Category");
                cells["B1"].PutValue("Value");
                cells["A2"].PutValue("A");
                cells["B2"].PutValue(10);
                cells["A3"].PutValue("B");
                cells["B3"].PutValue(20);
                cells["A4"].PutValue("C");
                cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                // Parameters: chart type, upper left row, upper left column, lower right row, lower right column
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // -------------------------------------------------
                // Step 1: Use automatic positioning for the legend
                // -------------------------------------------------
                Legend legend = chart.Legend;

                // By default the legend is positioned automatically.
                // Setting the Position property without specifying X/Y ratios lets Aspose.Cells handle placement.
                legend.Position = LegendPositionType.Top; // any position works; automatic placement is still applied
                Console.WriteLine("Legend set to automatic positioning.");

                // -------------------------------------------------
                // Step 2: Override legend position to a fixed top‑left location
                // -------------------------------------------------
                // Disable any automatic sizing (optional, keeps size unchanged)
                legend.IsAutomaticSize = false;

                // Set the legend to the top position (required before adjusting ratios)
                legend.Position = LegendPositionType.Top;

                // Move the legend to the exact top‑left corner of the chart area.
                // XRatioToChart and YRatioToChart are fractions of the chart area (0.0 = left/top edge).
                legend.XRatioToChart = 0.0; // left edge
                legend.YRatioToChart = 0.0; // top edge

                Console.WriteLine("Legend overridden to fixed top‑left location (XRatio=0, YRatio=0).");

                // Save the workbook to verify the result
                string outputPath = "ChartWithCustomLegend.xlsx";

                // Ensure the directory exists before saving
                try
                {
                    string directory = Path.GetDirectoryName(outputPath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine("Error saving workbook:");
                    Console.WriteLine(saveEx.Message);
                }
            }
            catch (Exception ex)
            {
                // Report any unexpected errors
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
