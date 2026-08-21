// Title: Export an AspNet.Cells line chart to scalable SVG using C#
// Description: Demonstrates how to create a workbook, add a line chart, configure SvgImageOptions (FitToViewPort, CSS prefix, WOFF font embedding), and save the chart as an SVG file with Aspose.Cells in C#.
// Keywords: Aspose.Cells SVG export | C# chart to SVG | FitToViewPort Aspose | SVG CSS prefix | embed WOFF font SVG | line chart SVG export .NET | scalable vector chart Aspose
// Common Searches: export Aspose.Cells chart as SVG C# | SvgImageOptions FitToViewPort example | how to add CSS prefix to SVG chart Aspose | embed fonts in SVG when exporting chart | scalable SVG chart for web Aspose.Cells
// Developer Intent: Generate a high‑quality SVG file from an Aspose.Cells chart with options that ensure scalability and style isolation.
// Use Cases: Integrate responsive SVG charts into web dashboards without rasterization artifacts. | Prevent CSS conflicts by applying a custom prefix to chart elements in the exported SVG. | Maintain consistent typography across browsers by embedding WOFF fonts directly in the SVG.
// AI Prompts: Provide C# code to export a column chart from Aspose.Cells to SVG with a transparent background and custom dimensions. | Show how to iterate through all charts in a workbook and save each as a uniquely named SVG file using Aspose.Cells. | Explain how to adjust the SVG viewport after export to fit a specific container size in a web page.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgExport
{
    // Demonstrates how to create a workbook, add a line chart, configure SvgImageOptions (FitToViewPort, CSS prefix, WOFF font embedding), and save the chart as an SVG file with Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Month");
                worksheet.Cells["A2"].PutValue("Jan");
                worksheet.Cells["A3"].PutValue("Feb");
                worksheet.Cells["A4"].PutValue("Mar");

                worksheet.Cells["B1"].PutValue("Sales");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(210);
                worksheet.Cells["B4"].PutValue(150);

                // Add a line chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories
                chart.Title.Text = "Quarterly Sales";

                // Configure SVG rendering options
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    FitToViewPort = true,                  // Make SVG fit the viewport for scalability
                    CssPrefix = "chart-",                  // Optional CSS prefix
                    EmbeddedFontType = SvgEmbeddedFontType.Woff // Optional font embedding
                };

                // Export the chart as an SVG file
                string outputPath = "QuarterlySalesChart.svg";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save chart to SVG
                chart.ToImage(outputPath, svgOptions);
                Console.WriteLine($"Chart exported to SVG: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
