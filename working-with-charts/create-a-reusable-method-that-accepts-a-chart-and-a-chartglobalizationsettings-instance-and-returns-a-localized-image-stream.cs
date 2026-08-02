// Title: C# helper method to render an Aspose.Cells chart with custom ChartGlobalizationSettings to a PNG MemoryStream
// Description: A static C# utility that accepts an Aspose.Cells Chart and a ChartGlobalizationSettings instance, applies the settings to the chart's workbook, renders the chart as a high‑resolution PNG into a MemoryStream, resets the stream position, and returns the stream for further use.
// Keywords: Aspose.Cells | ChartGlobalizationSettings | C# chart rendering | PNG MemoryStream | chart localization | export chart image | SettableChartGlobalizationSettings | high resolution chart image | Aspose.Cells ImageOrPrintOptions | chart to image stream
// Common Searches: Aspose.Cells render chart with ChartGlobalizationSettings | C# export chart to PNG memory stream | localized chart image Aspose.Cells .NET | how to apply SettableChartGlobalizationSettings to a chart | generate chart image stream without saving file
// Developer Intent: Create a PNG image stream of a chart that incorporates specified localization settings.
// Use Cases: Produce language‑specific chart images for multi‑region dashboards. | Return chart graphics directly from a Web API without creating temporary files. | Batch render charts with different locales before embedding them into PDFs or reports.
// AI Prompts: Write a C# method that takes an Aspose.Cells Chart and a ChartGlobalizationSettings object and returns a MemoryStream containing a PNG of the localized chart. | Show how to use SettableChartGlobalizationSettings to change a chart title and series name, then call the helper to get a localized image stream. | Demonstrate calling GetLocalizedChartImage from an ASP.NET Core controller and returning the result as a FileResult with the correct content type.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace ChartLocalizationDemo
{
    // A static C# utility that accepts an Aspose.Cells Chart and a ChartGlobalizationSettings instance, applies the settings to the chart's workbook, renders the chart as a high‑resolution PNG into a MemoryStream, resets the stream position, and returns the stream for further use.
    public static class ChartLocalizationHelper
    {
        /// <param name="chart">The chart to be rendered.</param>
        /// <param name="globalizationSettings">Custom globalization settings for the chart.</param>
        /// <returns>A MemoryStream containing the chart image (PNG format).</returns>
        public static MemoryStream GetLocalizedChartImage(Chart chart, ChartGlobalizationSettings globalizationSettings)
        {
            // Apply custom globalization settings to the workbook that owns the chart.
            chart.Worksheet.Workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = globalizationSettings
            };

            // Prepare a memory stream to receive the image data.
            var imageStream = new MemoryStream();

            // Configure image rendering options. PNG is the default format.
            var options = new ImageOrPrintOptions
            {
                // ImageFormat property is optional; default is PNG.
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Render the chart into the stream using the specified options.
            chart.ToImage(imageStream, options);

            // Reset the stream position so that callers can read from the beginning.
            imageStream.Position = 0;

            return imageStream;
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet.
                var wb = new Workbook();
                var ws = wb.Worksheets[0];

                // Populate worksheet with sample data.
                ws.Cells["A1"].PutValue("Category");
                ws.Cells["B1"].PutValue("Value");
                ws.Cells["A2"].PutValue("A");
                ws.Cells["A3"].PutValue("B");
                ws.Cells["A4"].PutValue("C");
                ws.Cells["A5"].PutValue("D");
                ws.Cells["B2"].PutValue(10);
                ws.Cells["B3"].PutValue(20);
                ws.Cells["B4"].PutValue(30);
                ws.Cells["B5"].PutValue(40);

                // Add a column chart.
                int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart myChart = ws.Charts[chartIdx];
                myChart.NSeries.Add("B2:B5", true);
                myChart.NSeries.CategoryData = "A2:A5";

                // Create custom globalization settings.
                var customSettings = new SettableChartGlobalizationSettings();
                customSettings.SetChartTitleName("自定义标题");
                customSettings.SetSeriesName("自定义系列");

                // Generate the localized image stream.
                MemoryStream localizedImage = ChartLocalizationHelper.GetLocalizedChartImage(myChart, customSettings);

                // Save the image to a file.
                string outputPath = "LocalizedChart.png";
                using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    localizedImage.CopyTo(file);
                }

                Console.WriteLine($"Chart image saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
