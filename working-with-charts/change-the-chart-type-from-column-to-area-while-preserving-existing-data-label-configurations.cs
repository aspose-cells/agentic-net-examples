// Title: Convert a column chart to an area chart while preserving series data label settings using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook, records each series' data label properties, switches the first chart's type to Area, and reapplies the saved label settings before saving. | Provide an Aspose.Cells script that changes a column chart to an area chart without losing any data label formatting for all series.
// Common Searches: Aspose.Cells change chart type from column to area without losing data label formatting | C# preserve series data labels when converting Excel chart type | how to keep data label settings after switching chart type in Aspose.Cells | convert first worksheet chart to area chart while retaining label options using .NET
// Tags: Aspose.Cells chart type conversion area | preserve series data labels Aspose.Cells | C# change Excel chart from column to area | data label settings retention .NET Excel chart | modify chart type programmatically Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an existing workbook, captures each series' data label configuration, changes the first chart's type to Area, restores the saved label settings, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart on the sheet
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the worksheet.");
                return;
            }

            Chart chart = sheet.Charts[0];

            // Preserve series‑level data label configurations
            var seriesLabelConfigs = new List<SeriesLabelConfig>();
            foreach (Series series in chart.NSeries)
            {
                SeriesLabelConfig cfg = new SeriesLabelConfig
                {
                    ShowValue = series.DataLabels.ShowValue,
                    ShowCategoryName = series.DataLabels.ShowCategoryName,
                    ShowSeriesName = series.DataLabels.ShowSeriesName,
                    ShowPercentage = series.DataLabels.ShowPercentage,
                    ShowBubbleSize = series.DataLabels.ShowBubbleSize,
                    NumberFormat = series.DataLabels.NumberFormat
                };
                seriesLabelConfigs.Add(cfg);
            }

            // Change the chart type to Area
            chart.Type = ChartType.Area;

            // Restore series‑level data label configurations
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                Series series = chart.NSeries[i];
                SeriesLabelConfig cfg = seriesLabelConfigs[i];

                series.DataLabels.ShowValue = cfg.ShowValue;
                series.DataLabels.ShowCategoryName = cfg.ShowCategoryName;
                series.DataLabels.ShowSeriesName = cfg.ShowSeriesName;
                series.DataLabels.ShowPercentage = cfg.ShowPercentage;
                series.DataLabels.ShowBubbleSize = cfg.ShowBubbleSize;
                series.DataLabels.NumberFormat = cfg.NumberFormat;
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper class to store a series' data label configuration
    private class SeriesLabelConfig
    {
        public bool ShowValue { get; set; }
        public bool ShowCategoryName { get; set; }
        public bool ShowSeriesName { get; set; }
        public bool ShowPercentage { get; set; }
        public bool ShowBubbleSize { get; set; }
        public string NumberFormat { get; set; } = string.Empty;
    }
}
