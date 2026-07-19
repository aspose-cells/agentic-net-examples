// Title: C# helper to export an Aspose.Cells chart with ChartGlobalizationSettings as a PNG MemoryStream
// Description: A static method that receives a Chart object and a ChartGlobalizationSettings instance, applies the settings to the chart's workbook, configures PNG output at 300 dpi, renders the chart into a MemoryStream, and returns the stream ready for saving, API response, or further processing.
// Keywords: Aspose.Cells chart export | ChartGlobalizationSettings | C# PNG MemoryStream | chart ToImage Aspose | localized chart image | ImageOrPrintOptions PNG | render chart to stream
// Common Searches: Aspose.Cells render chart with custom globalization to PNG | How to get a chart image as MemoryStream in C# | Apply ChartGlobalizationSettings before exporting chart | Export Aspose chart to PNG without saving a file | Generate localized chart image programmatically
// Developer Intent: Produce a PNG image stream of a chart after applying user‑defined globalization settings.
// Use Cases: Create language‑specific chart graphics for PDF or HTML reports. | Return chart PNG data directly from a web API, avoiding temporary files. | Cache in‑memory chart images for dashboards while preserving localized titles and series names.
// AI Prompts: Write a C# method that takes an Aspose.Cells Chart and a ChartGlobalizationSettings object, applies the settings, and returns a 300 dpi PNG MemoryStream. | Show how to configure ImageOrPrintOptions for PNG output and use Chart.ToImage to write the result into a MemoryStream. | Provide robust null‑argument checks and stream positioning best practices when exporting a chart to a MemoryStream with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;   // Required for ImageType enum

// A static method that receives a Chart object and a ChartGlobalizationSettings instance, applies the settings to the chart's workbook, configures PNG output at 300 dpi, renders the chart into a MemoryStream, and returns the stream ready for saving, API response, or further processing.
public static class ChartLocalizationHelper
{
    /// <param name="chart">The chart to be rendered.</param>
    /// <param name="globalizationSettings">Custom globalization settings for the chart.</param>
    /// <returns>A MemoryStream containing the chart image (PNG format).</returns>
    public static MemoryStream GetLocalizedChartImage(Chart chart, ChartGlobalizationSettings globalizationSettings)
    {
        if (chart == null) throw new ArgumentNullException(nameof(chart));
        if (globalizationSettings == null) throw new ArgumentNullException(nameof(globalizationSettings));

        // Apply the custom globalization settings to the workbook that owns the chart.
        Workbook workbook = chart.Worksheet.Workbook;
        workbook.Settings.GlobalizationSettings = new GlobalizationSettings
        {
            ChartSettings = globalizationSettings
        };

        // Prepare image options – PNG format with a reasonable resolution.
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,
            HorizontalResolution = 300,
            VerticalResolution = 300
        };

        // Render the chart to a memory stream.
        MemoryStream imageStream = new MemoryStream();
        chart.ToImage(imageStream, options);
        imageStream.Position = 0; // Reset position for downstream consumers.

        return imageStream;
    }
}

public class Example
{
    public static void Run()
    {
        try
        {
            // Create a workbook and a simple chart.
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);

            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = ws.Charts[chartIdx];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Create custom globalization settings (example: change series name and title).
            var customSettings = new SettableChartGlobalizationSettings();
            customSettings.SetSeriesName("Custom Series");
            customSettings.SetChartTitleName("Localized Chart");

            // Generate the localized image stream.
            MemoryStream imgStream = ChartLocalizationHelper.GetLocalizedChartImage(chart, customSettings);

            // Save the stream to a file.
            string outputPath = "LocalizedChart.png";
            using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                imgStream.CopyTo(file);
            }

            Console.WriteLine($"Localized chart image saved as '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        Example.Run();
    }
}
