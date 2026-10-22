// Title: How to create a reusable C# method that renders an Aspose.Cells chart to a localized PNG MemoryStream using ChartGlobalizationSettings
// AI Prompts: Generate a static C# method named GetLocalizedChartImage that accepts a Chart object and a ChartGlobalizationSettings instance, applies the globalization settings to ImageOrPrintOptions via reflection when available, renders the chart to a temporary PNG file, returns the image as a MemoryStream, and deletes the temporary file. | Write C# sample code that loads a workbook, retrieves the first chart, creates a ChartGlobalizationSettings object, calls GetLocalizedChartImage, and saves the resulting stream to a PNG file while handling exceptions and disposing resources.
// Common Searches: C# Aspose.Cells render chart to PNG with culture-specific settings | How to use ChartGlobalizationSettings when exporting a chart image in Aspose.Cells | Get chart image as MemoryStream without leaving temporary files in C# | Reflection example for setting ImageOrPrintOptions.ChartGlobalizationSettings property | Reusable method to localize Aspose.Cells chart output as PNG stream
// Tags: render chart to PNG stream Aspose.Cells | apply ChartGlobalizationSettings via reflection | temporary file cleanup after chart rendering | localized chart image generation C# | chart ToImage with culture settings

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // Provides a static GetLocalizedChartImage method that takes an Aspose.Cells Chart and an optional ChartGlobalizationSettings object, sets the globalization settings on ImageOrPrintOptions using reflection when the property exists, renders the chart to a temporary PNG file, loads the file into a MemoryStream, deletes the temporary file, and returns the stream for further use.
    public static class ChartHelper
    {
        /// <param name="chart">The Aspose.Cells chart to render.</param>
        /// <param name="globalizationSettings">Culture‑specific settings for rendering (may be null).</param>
        /// <returns>A stream containing the PNG image. The caller must dispose the stream.</returns>
        public static Stream GetLocalizedChartImage(Chart chart, ChartGlobalizationSettings globalizationSettings)
        {
            if (chart == null) throw new ArgumentNullException(nameof(chart));

            // Prepare rendering options.
            var options = new ImageOrPrintOptions();

            // Attempt to set globalization settings via reflection (property may not exist in some versions).
            if (globalizationSettings != null)
            {
                try
                {
                    var prop = typeof(ImageOrPrintOptions).GetProperty("ChartGlobalizationSettings");
                    if (prop != null && prop.CanWrite)
                    {
                        prop.SetValue(options, globalizationSettings);
                    }
                }
                catch
                {
                    // Ignore if reflection fails.
                }
            }

            // Render chart to a temporary PNG file, then load it into a memory stream.
            string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".png");
            try
            {
                chart.ToImage(tempPath, options);
                var imageBytes = File.ReadAllBytes(tempPath);
                return new MemoryStream(imageBytes);
            }
            finally
            {
                // Clean up the temporary file.
                if (File.Exists(tempPath))
                {
                    try { File.Delete(tempPath); } catch { /* ignore cleanup errors */ }
                }
            }
        }
    }

    internal class Program
    {
        private static void Main()
        {
            try
            {
                const string workbookPath = "sample.xlsx";
                const string outputImagePath = "chart.png";

                // Verify that the input workbook exists.
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Input file not found: {workbookPath}");
                    return;
                }

                // Load the workbook.
                var workbook = new Workbook(workbookPath);

                // Assume the first worksheet contains at least one chart.
                var sheet = workbook.Worksheets[0];
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the first worksheet.");
                    return;
                }

                // Get the first chart.
                var chart = sheet.Charts[0];

                // Prepare globalization settings (if needed). No culture property is set to keep compatibility.
                var globalizationSettings = new ChartGlobalizationSettings();

                // Render chart to PNG stream.
                using (var pngStream = ChartHelper.GetLocalizedChartImage(chart, globalizationSettings))
                using (var fileStream = new FileStream(outputImagePath, FileMode.Create, FileAccess.Write))
                {
                    pngStream.CopyTo(fileStream);
                }

                Console.WriteLine($"Chart image saved to {outputImagePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
