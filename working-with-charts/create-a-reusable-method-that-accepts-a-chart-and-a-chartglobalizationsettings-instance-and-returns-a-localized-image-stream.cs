// Title: C# helper to render an Aspose.Cells chart with ChartGlobalizationSettings into a PNG MemoryStream
// Description: Provides a reusable method that validates a Chart object, applies a ChartGlobalizationSettings instance to the owning workbook, configures 300 DPI PNG rendering via ImageOrPrintOptions, and returns the chart image as a MemoryStream. Includes a complete example that creates sample data, builds a column chart, sets localized series and title names, and saves the resulting image.
// Keywords: Aspose.Cells chart localization | ChartGlobalizationSettings C# | render chart to PNG stream | Aspose.Cells ImageOrPrintOptions | C# memory stream chart image | .NET Excel chart export | localized chart thumbnail
// Common Searches: Aspose.Cells apply ChartGlobalizationSettings to a chart | C# render Aspose.Cells chart as PNG stream | How to export an Aspose.Cells chart to MemoryStream | Generate localized chart image with Aspose.Cells for .NET | ChartGlobalizationSettings example code
// Developer Intent: Create a PNG image stream of an Excel chart that reflects custom globalization (locale‑specific titles, series names, etc.) without writing intermediate files.
// Use Cases: Produce language‑specific chart images for multi‑region reporting dashboards. | Serve chart thumbnails directly from a web API using a MemoryStream response. | Embed localized chart graphics into PDFs or Word documents generated on the server. | Cache chart images per locale to improve performance in international applications.
// AI Prompts: Show how to call GetLocalizedChartImage with French ChartGlobalizationSettings and write the result to a file. | Extend GetLocalizedChartImage to accept an output format (PNG, JPEG) and a custom DPI value. | Write unit tests that verify ArgumentNullException for null chart or settings and that InvalidOperationException wraps rendering errors. | Explain how to reuse the helper in an ASP.NET Core controller that returns FileStreamResult.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

// Provides a reusable method that validates a Chart object, applies a ChartGlobalizationSettings instance to the owning workbook, configures 300 DPI PNG rendering via ImageOrPrintOptions, and returns the chart image as a MemoryStream. Includes a complete example that creates sample data, builds a column chart, sets localized series and title names, and saves the resulting image.
public static class ChartLocalizationHelper
{
    /// <param name="chart">The chart to be rendered.</param>
    /// <param name="globalizationSettings">Custom globalization settings for the chart.</param>
    /// <returns>A MemoryStream containing the chart image (PNG format).</returns>
    public static MemoryStream GetLocalizedChartImage(Chart chart, ChartGlobalizationSettings globalizationSettings)
    {
        if (chart == null) throw new ArgumentNullException(nameof(chart));
        if (globalizationSettings == null) throw new ArgumentNullException(nameof(globalizationSettings));

        try
        {
            // Apply the custom globalization settings to the workbook that owns the chart
            Workbook workbook = chart.Worksheet.Workbook;
            workbook.Settings.GlobalizationSettings = new GlobalizationSettings
            {
                ChartSettings = globalizationSettings
            };

            // Set image rendering options (PNG, 300 DPI)
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                // Default format is PNG; explicit setting omitted to avoid API mismatch
                HorizontalResolution = 300,
                VerticalResolution = 300
            };

            // Render the chart into a memory stream
            MemoryStream imageStream = new MemoryStream();
            chart.ToImage(imageStream, options);
            imageStream.Position = 0; // Reset for downstream consumers

            return imageStream;
        }
        catch (Exception ex)
        {
            // Wrap and rethrow to preserve stack trace while providing context
            throw new InvalidOperationException("Failed to generate localized chart image.", ex);
        }
    }
}

public class Example
{
    public static void Run()
    {
        try
        {
            // Create a workbook and populate sample data
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);

            // Add a column chart
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = ws.Charts[chartIdx];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Create custom globalization settings (e.g., custom series and title names)
            var customSettings = new SettableChartGlobalizationSettings();
            customSettings.SetSeriesName("Custom Series");
            customSettings.SetChartTitleName("Localized Chart");

            // Generate the localized chart image
            MemoryStream imgStream = ChartLocalizationHelper.GetLocalizedChartImage(chart, customSettings);

            // Save the image stream to a file
            string outputPath = "LocalizedChart.png";
            using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                imgStream.CopyTo(file);
            }

            Console.WriteLine($"Localized chart image saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main()
    {
        Example.Run();
    }
}
