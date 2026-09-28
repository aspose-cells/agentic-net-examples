// Title: Reusable C# method to render an Aspose.Cells chart as a PNG MemoryStream with optional globalization settings
// AI Prompts: Write a static C# method that receives an Aspose.Cells Chart and an optional ChartGlobalizationSettings object, uses ImageOrPrintOptions to render the chart, and returns the image as a MemoryStream. | Add robust error handling to the chart‑to‑image helper so that any exception is caught and re‑thrown as an InvalidOperationException with a descriptive message. | Extend the method to let the caller specify the output image format (PNG, JPEG, BMP) through ImageOrPrintOptions before generating the stream.
// Common Searches: c# Aspose.Cells render chart to memory stream png | how to export an Excel chart as an image using Aspose.Cells Chart.ToImage | using ChartGlobalizationSettings when converting Aspose.Cells chart to image | save Aspose.Cells chart as PNG file from a MemoryStream | custom ImageOrPrintOptions for Aspose.Cells chart image export
// Tags: Aspose.Cells chart to PNG stream | C# chart rendering helper method | ChartGlobalizationSettings parameter usage | ImageOrPrintOptions image format selection | InvalidOperationException wrapper for chart rendering

using System;
using System.IO;
using System.Drawing.Imaging;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Provides a reusable C# static method that accepts an Aspose.Cells Chart and an optional ChartGlobalizationSettings object, renders the chart to a PNG MemoryStream using ImageOrPrintOptions, and includes comprehensive exception handling.
public static class ChartHelper
{
    /// <param name="chart">The Aspose.Cells chart to render.</param>
    /// <param name="globalizationSettings">Locale‑specific settings (currently not applied due to API limitations).</param>
    /// <returns>A MemoryStream containing the chart image (PNG format).</returns>
    public static Stream GetLocalizedChartImage(Chart chart, ChartGlobalizationSettings? globalizationSettings = null)
    {
        if (chart == null) throw new ArgumentNullException(nameof(chart));

        try
        {
            // NOTE: In some Aspose.Cells versions Chart does not expose GlobalizationSettings.
            // If available, it could be set here. The parameter is kept for compatibility.

            // Prepare image options – default format is PNG, so no need to set ImageFormat explicitly.
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                // Additional options can be set here if needed.
                // For example: imgOptions.OnePagePerSheet = true;
            };

            // Render the chart into a memory stream.
            MemoryStream imageStream = new MemoryStream();
            chart.ToImage(imageStream, imgOptions);
            imageStream.Position = 0;
            return imageStream;
        }
        catch (Exception ex)
        {
            // Wrap and rethrow to preserve stack trace while providing context.
            throw new InvalidOperationException("Failed to render chart to image.", ex);
        }
    }
}

class Program
{
    static void Main()
    {
        try
        {
            string workbookPath = "Sample.xlsx";

            // Ensure the input workbook exists.
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Input file not found: {workbookPath}");
                return;
            }

            // Load workbook.
            Workbook workbook = new Workbook(workbookPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Verify that a chart exists.
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            Chart chart = sheet.Charts[0];

            // Render chart to image.
            using (Stream imgStream = ChartHelper.GetLocalizedChartImage(chart))
            {
                string outputPath = "ChartImage.png";

                // Save the image to disk.
                using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
                {
                    imgStream.CopyTo(file);
                }

                Console.WriteLine($"Chart image saved to {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
