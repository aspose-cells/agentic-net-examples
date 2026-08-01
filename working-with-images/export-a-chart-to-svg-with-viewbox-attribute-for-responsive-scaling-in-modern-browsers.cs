// Title: Export Aspose.Cells Chart to SVG with viewBox for Responsive Scaling (C#)
// Description: Creates a workbook, adds sample data, builds a column chart, and uses Aspose.Cells SvgImageOptions (FitToViewPort = true) to save the chart as an SVG file that contains a viewBox attribute, enabling responsive scaling in modern browsers.
// Keywords: Aspose.Cells | C# | export chart to SVG | viewBox attribute | FitToViewPort | responsive SVG | SvgImageOptions | Excel chart to SVG | scalable vector graphics | column chart
// Common Searches: Aspose.Cells export chart as SVG C# | How to add viewBox to SVG chart using Aspose.Cells | FitToViewPort option example Aspose.Cells | Responsive SVG chart from Excel workbook | C# code to save chart to SVG with viewBox | Aspose.Cells SvgImageOptions responsive scaling
// Developer Intent: Generate an SVG file from an Aspose.Cells chart that includes a viewBox attribute for responsive display.
// Use Cases: Embed scalable SVG charts in web dashboards that automatically adapt to container size. | Create high‑resolution graphics for email newsletters or PDF reports without pixelation. | Provide vector‑based visualizations in mobile apps that adjust to different screen orientations.
// AI Prompts: Show C# code to export multiple Aspose.Cells charts to SVG files with viewBox enabled. | Explain how to customize the CSS prefix and embed fonts when exporting a chart to SVG with Aspose.Cells. | Demonstrate adjusting the output dimensions while preserving the viewBox for responsive behavior.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgExport
{
    // Creates a workbook, adds sample data, builds a column chart, and uses Aspose.Cells SvgImageOptions (FitToViewPort = true) to save the chart as an SVG file that contains a viewBox attribute, enabling responsive scaling in modern browsers.
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

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories
                chart.Title.Text = "Quarterly Sales";

                // Configure SVG rendering options (no ImageFormat property needed)
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    FitToViewPort = true,            // Enable viewBox for responsive scaling
                    CssPrefix = "chart-",            // Optional: CSS prefix
                    EmbeddedFontType = SvgEmbeddedFontType.Woff // Optional: embed fonts
                };

                // Export the chart to an SVG file with viewBox attribute
                string outputPath = "QuarterlySalesChart.svg";
                chart.ToImage(outputPath, svgOptions);

                Console.WriteLine($"Chart exported to SVG with viewBox: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
