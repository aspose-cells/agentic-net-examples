// Title: How to set individual legend entry font colors by series index while keeping legend background transparent using Aspose.Cells for .NET
// AI Prompts: Iterate through a chart's LegendEntries collection and assign Font.Color from a predefined Color[] based on the entry index in C# with Aspose.Cells. | Load or create a workbook, ensure a chart exists, and customize legend entry fonts without changing the entry fill transparency using the Aspose.Cells API.
// Common Searches: Aspose.Cells set legend entry font color per series in C# | keep legend background transparent Aspose.Cells chart | apply different colors to Excel chart legend items using Aspose.Cells .NET | change legend font colors based on series index Aspose.Cells example | C# Aspose.Cells legend entry styling without fill change
// Tags: legend entry color styling Aspose.Cells | series-index based legend formatting .NET | transparent legend fill Aspose.Cells | chart legend color array C# | Aspose.Cells legend customization

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads or creates a workbook, obtains the first chart (or adds one if missing), defines a Color array, loops through the chart's legend entries, sets each entry's Font.Color according to its index, leaves the legend entry background transparent, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Load existing workbook or create a new one if the file is missing.
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // default workbook with one worksheet
                }

                // Get the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Obtain an existing chart or create a new placeholder chart.
                Chart chart;
                if (sheet.Charts.Count > 0)
                {
                    chart = sheet.Charts[0];
                }
                else
                {
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                    chart = sheet.Charts[chartIndex];

                    // Add a dummy series so legend entries exist.
                    int seriesIndex = chart.NSeries.Add("A1:A5", true);
                    var series = chart.NSeries[seriesIndex];
                    series.Name = "Series 1";
                }

                // Define colors for legend fonts.
                Color[] legendColors = new Color[]
                {
                    Color.Red,
                    Color.Green,
                    Color.Blue,
                    Color.Orange,
                    Color.Purple
                };

                // Apply font colors to each legend entry if legend exists.
                if (chart.Legend != null && chart.Legend.LegendEntries.Count > 0)
                {
                    int entryCount = chart.Legend.LegendEntries.Count;
                    for (int i = 0; i < entryCount; i++)
                    {
                        LegendEntry entry = chart.Legend.LegendEntries[i];
                        Color fontColor = legendColors[i % legendColors.Length];
                        entry.Font.Color = fontColor;
                        // Background of legend entries is transparent by default.
                    }
                }

                // Save the workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
