// Title: C# example: Convert a column chart to an area chart while keeping data label settings with Aspose.Cells
// Common Searches: Aspose.Cells C# change column chart to area chart keep data labels | how to retain data label configuration when switching chart type in Aspose.Cells | C# code to preserve series data labels after changing chart type in Excel | convert Excel column chart to area chart using Aspose.Cells .NET | sample program to modify chart type without losing label settings Aspose.Cells
// Tags: aspose.cells area chart conversion | c# retain chart data label settings | excel chart type modification aspose | preserve series labels aspose.cells | chart series data label reuse c#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChangeChartTypeExample
{
    // Helper class to store data label settings for a series
    // The program loads an existing Excel workbook, extracts each series' data label settings from the first chart, changes the chart type from Column to Area, reapplies the saved label configurations, and saves the updated workbook to a new file.
    internal class DataLabelSettings
    {
        public bool ShowValue { get; set; }
        public bool ShowCategoryName { get; set; }
        public bool ShowSeriesName { get; set; }
        public bool ShowPercentage { get; set; }
        public bool ShowBubbleSize { get; set; }
        public bool ShowLegendKey { get; set; }
        public LabelPositionType Position { get; set; } // Correct enum type
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (adjust as needed)
            string inputFile = "input.xlsx";
            string outputFile = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file '{inputFile}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFile);
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one chart in the worksheet
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the first worksheet.");
                    return;
                }

                // Get the first chart (assumed to be a column chart)
                Chart chart = worksheet.Charts[0];

                // Preserve existing data label configurations for each series
                List<DataLabelSettings> labelSettings = new List<DataLabelSettings>();
                for (int i = 0; i < chart.NSeries.Count; i++)
                {
                    Series series = chart.NSeries[i];
                    DataLabels dl = series.DataLabels;

                    DataLabelSettings settings = new DataLabelSettings
                    {
                        ShowValue = dl.ShowValue,
                        ShowCategoryName = dl.ShowCategoryName,
                        ShowSeriesName = dl.ShowSeriesName,
                        ShowPercentage = dl.ShowPercentage,
                        ShowBubbleSize = dl.ShowBubbleSize,
                        ShowLegendKey = dl.ShowLegendKey,
                        Position = dl.Position
                    };

                    labelSettings.Add(settings);
                }

                // Change the chart type from Column to Area
                chart.Type = ChartType.Area;

                // Reapply the preserved data label settings
                for (int i = 0; i < chart.NSeries.Count && i < labelSettings.Count; i++)
                {
                    Series series = chart.NSeries[i];
                    DataLabels dl = series.DataLabels;
                    DataLabelSettings settings = labelSettings[i];

                    dl.ShowValue = settings.ShowValue;
                    dl.ShowCategoryName = settings.ShowCategoryName;
                    dl.ShowSeriesName = settings.ShowSeriesName;
                    dl.ShowPercentage = settings.ShowPercentage;
                    dl.ShowBubbleSize = settings.ShowBubbleSize;
                    dl.ShowLegendKey = settings.ShowLegendKey;
                    dl.Position = settings.Position;
                }

                // Save the modified workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Chart type changed to Area and saved to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
