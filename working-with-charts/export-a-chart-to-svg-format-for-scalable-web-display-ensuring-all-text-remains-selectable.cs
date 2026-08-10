// Title: Export Aspose.Cells Chart to SVG with Selectable Text in C#
// Description: Creates a workbook, adds a column chart from sample data, configures SvgImageOptions (viewport fit, CSS prefix, WOFF font embedding) and saves the chart as an SVG file where all text remains selectable in browsers.
// Keywords: Aspose.Cells | C# chart export | SVG output | selectable text SVG | SvgImageOptions | EmbeddedFontType Woff | FitToViewPort | chart CSS prefix | Aspose.Cells ToImage | web‑ready chart
// Common Searches: Aspose.Cells export chart to SVG C# | how to keep text selectable when saving chart as SVG | C# Aspose.Cells SVG rendering options | embed fonts in SVG chart Aspose.Cells | fit chart to viewport SVG Aspose
// Developer Intent: Generate an SVG file from an Aspose.Cells chart while preserving selectable text.
// Use Cases: Render responsive SVG charts for web dashboards. | Embed WOFF fonts in SVG to maintain text accessibility. | Apply a custom CSS prefix to style exported chart elements consistently. | Export workbook charts without rasterizing text for SEO‑friendly graphics.
// AI Prompts: Show C# code that uses Aspose.Cells to export a chart to SVG with embedded WOFF fonts for selectable text. | Explain how to configure SvgImageOptions to set FitToViewPort and a custom CSS prefix when saving a chart as SVG. | Provide error‑handling best practices for chart-to‑SVG conversion in a .NET application.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsSvgExport
{
    // Creates a workbook, adds a column chart from sample data, configures SvgImageOptions (viewport fit, CSS prefix, WOFF font embedding) and saves the chart as an SVG file where all text remains selectable in browsers.
    public class ExportChartToSvg
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("Apple");
                worksheet.Cells["A3"].PutValue("Orange");
                worksheet.Cells["A4"].PutValue("Banana");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(80);
                worksheet.Cells["B4"].PutValue(150);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";     // Categories
                chart.Title.Text = "Fruit Sales";

                // Configure SVG rendering options
                SvgImageOptions svgOptions = new SvgImageOptions();
                svgOptions.ImageType = ImageType.Svg;               // Ensure SVG output
                svgOptions.FitToViewPort = true;                   // Fit to viewport
                svgOptions.CssPrefix = "chart-";                   // Optional CSS prefix
                svgOptions.EmbeddedFontType = SvgEmbeddedFontType.Woff; // Embed font for selectable text

                // Export the chart as an SVG file; text remains selectable because we use SVG format
                string outputPath = "FruitSalesChart.svg";
                chart.ToImage(outputPath, svgOptions);

                Console.WriteLine($"Chart exported to SVG: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point required by the project
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportChartToSvg.Run();
        }
    }
}
