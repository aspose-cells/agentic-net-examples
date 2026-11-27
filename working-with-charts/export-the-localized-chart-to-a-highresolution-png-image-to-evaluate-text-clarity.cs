// Title: Export a Japanese‑localized column chart to a high‑resolution PNG file with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export column chart with Japanese text to 300 DPI PNG | How to set chart title font for Japanese characters in Aspose.Cells | Render high‑resolution chart image using ImageOrPrintOptions in .NET | Save localized chart as PNG with specific DPI using Aspose.Cells | Aspose.Cells chart ToImage high DPI PNG example
// Tags: export chart to PNG Aspose.Cells | high DPI chart rendering ImageOrPrintOptions | Japanese font support chart title Aspose.Cells | column chart image generation C# | localized chart image export .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace ExportLocalizedChartHighRes
{
    // The sample creates a workbook with Japanese category labels, builds a column chart with a Japanese title and legend using the MS Gothic font, configures ImageOrPrintOptions for 300 DPI PNG output, and saves the chart as a high‑resolution PNG file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and obtain the first worksheet.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -----------------------------------------------------------------
                // 2. Populate sample data that will be used by the chart.
                // -----------------------------------------------------------------
                // Header row
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // Sample data rows
                sheet.Cells["A2"].PutValue("製品A"); // Japanese text for "Product A"
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("製品B");
                sheet.Cells["B3"].PutValue(85);
                sheet.Cells["A4"].PutValue("製品C");
                sheet.Cells["B4"].PutValue(150);

                // -----------------------------------------------------------------
                // 3. Add a column chart that references the data range.
                // -----------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Define the data series and categories.
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Localize the chart title (e.g., Japanese).
                chart.Title.Text = "売上高"; // "Sales Amount" in Japanese
                chart.Title.Font.Name = "MS Gothic"; // Ensure a font that supports Japanese characters.
                chart.Title.Font.Size = 14;
                chart.Title.Font.IsBold = true;

                // Optional: set legend font to support localization.
                chart.Legend.Font.Name = "MS Gothic";
                chart.Legend.Font.Size = 12;

                // -----------------------------------------------------------------
                // 4. Configure high‑resolution image rendering options.
                // -----------------------------------------------------------------
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Png,
                    HorizontalResolution = 300, // 300 DPI horizontal
                    VerticalResolution = 300,   // 300 DPI vertical
                    // Ensure the chart is rendered with the same DPI settings.
                    // The following properties are optional but help keep the output size predictable.
                    OnePagePerSheet = true,
                    Transparent = false
                };

                // -----------------------------------------------------------------
                // 5. Export the chart to a PNG file with the specified DPI.
                // -----------------------------------------------------------------
                string outputPath = Path.Combine(Environment.CurrentDirectory, "localized_chart_highres.png");
                chart.ToImage(outputPath, imgOptions);

                Console.WriteLine($"Chart successfully exported to high‑resolution PNG:");
                Console.WriteLine(outputPath);
                Console.WriteLine($"Resolution: {imgOptions.HorizontalResolution} DPI (Horizontal), {imgOptions.VerticalResolution} DPI (Vertical)");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred while exporting the chart:");
                Console.Error.WriteLine(ex.Message);
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
