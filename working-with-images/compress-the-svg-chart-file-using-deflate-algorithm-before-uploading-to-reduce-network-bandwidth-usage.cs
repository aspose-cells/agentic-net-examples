// Title: Compress Aspose.Cells SVG Chart to .svgz with DeflateStream in C#
// Description: Creates a workbook, generates a column chart, renders it to SVG using Aspose.Cells, then compresses the SVG with .NET's DeflateStream to produce a .svgz file, lowering file size for faster network transfer.
// Keywords: Aspose.Cells SVG compression C# | DeflateStream .svgz | compress SVG chart .NET | reduce SVG size Aspose | upload compressed chart image
// Common Searches: Aspose.Cells compress SVG output | C# DeflateStream example for SVG | Save chart as .svgz using Aspose.Cells | How to shrink SVG file size before upload | Render Aspose.Cells chart to compressed SVG
// Developer Intent: Generate an SVG chart with Aspose.Cells and apply Deflate compression to create a smaller .svgz file for bandwidth‑efficient upload.
// Use Cases: Web API that returns chart images as compressed .svgz to minimize response payload. | Batch processing of reports where SVG charts are stored in a compressed format for archival. | Real‑time dashboards that serve compressed SVG charts to mobile clients with limited bandwidth.
// AI Prompts: Write C# code that creates an Aspose.Cells chart, exports it to SVG, and compresses the SVG with DeflateStream, handling missing files gracefully. | Explain the steps to configure SvgImageOptions for optimal SVG output before compression and how to decompress a .svgz back to SVG in .NET. | Suggest best practices for naming, caching, and serving .svgz files generated from Aspose.Cells charts in a web application.

using System;
using System.IO;
using System.IO.Compression;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgCompressionDemo
{
    // Creates a workbook, generates a column chart, renders it to SVG using Aspose.Cells, then compresses the SVG with .NET's DeflateStream to produce a .svgz file, lowering file size for faster network transfer.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["B4"].PutValue(25);

                // 2. Add a column chart
                int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";
                chart.Title.Text = "Sample Chart";

                // 3. Configure SVG rendering options
                SvgImageOptions svgOpts = new SvgImageOptions
                {
                    FitToViewPort = true,
                    CssPrefix = "demo-",
                    EmbeddedFontType = SvgEmbeddedFontType.Woff
                };

                // 4. Render the chart to an SVG file
                string svgPath = "chart.svg";
                chart.ToImage(svgPath, svgOpts);

                // 5. Read the generated SVG file into a byte array (ensure file exists)
                byte[] svgBytes = File.Exists(svgPath) ? File.ReadAllBytes(svgPath) : Array.Empty<byte>();

                // 6. Compress the SVG bytes using Deflate algorithm
                string compressedPath = "chart.svgz"; // .svgz is a common extension for deflated SVG
                using (FileStream outStream = new FileStream(compressedPath, FileMode.Create, FileAccess.Write))
                using (DeflateStream deflate = new DeflateStream(outStream, CompressionLevel.Optimal))
                {
                    if (svgBytes.Length > 0)
                    {
                        deflate.Write(svgBytes, 0, svgBytes.Length);
                    }
                }

                Console.WriteLine($"SVG chart saved to '{svgPath}'.");
                Console.WriteLine($"Compressed SVG saved to '{compressedPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
