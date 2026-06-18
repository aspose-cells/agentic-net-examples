using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgCompressionDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");

                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(210);
                sheet.Cells["B4"].PutValue(150);

                // 3. Add a line chart
                int chartIdx = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // 4. Configure SVG rendering options (no ImageFormat property)
                SvgImageOptions svgOpts = new SvgImageOptions
                {
                    FitToViewPort = true,        // Fit SVG to viewport
                    CssPrefix = "demo-",         // Optional CSS prefix
                    EmbeddedFontType = SvgEmbeddedFontType.Woff
                };

                // 5. Save the chart as an SVG file
                string svgPath = "chart.svg";
                chart.ToImage(svgPath, svgOpts);

                // 6. Compress the generated SVG using Deflate (produces .svgz)
                string compressedPath = "chart.svgz";

                if (File.Exists(svgPath))
                {
                    try
                    {
                        using (FileStream source = new FileStream(svgPath, FileMode.Open, FileAccess.Read))
                        using (FileStream dest = new FileStream(compressedPath, FileMode.Create, FileAccess.Write))
                        using (DeflateStream compression = new DeflateStream(dest, CompressionLevel.Optimal))
                        {
                            source.CopyTo(compression);
                        }

                        Console.WriteLine($"SVG chart saved to '{svgPath}'.");
                        Console.WriteLine($"Compressed SVG saved to '{compressedPath}'.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Compression error: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine($"Error: SVG file '{svgPath}' was not created.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}