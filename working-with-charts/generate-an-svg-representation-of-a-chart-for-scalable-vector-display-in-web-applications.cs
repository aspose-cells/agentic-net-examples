// Title: Export a Line Chart to SVG using Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills month‑sales data, adds a line chart, configures SvgImageOptions (fit to viewport, CSS prefix, WOFF font), and renders the chart directly to an SVG file while optionally saving the workbook.
// Keywords: Aspose.Cells | C# SVG chart export | line chart SVG | SvgImageOptions | Aspose.Cells chart rendering | export Excel chart to SVG | scalable vector graphics .NET | web dashboard SVG chart | chart to SVG Aspose
// Common Searches: Aspose.Cells export chart to SVG C# | How to render Excel chart as SVG using Aspose.Cells | SvgImageOptions example for chart in C# | Generate SVG line chart from workbook Aspose.Cells | Save Aspose.Cells chart as SVG file
// Developer Intent: The developer wants to generate a line chart from worksheet data and export it as an SVG file using Aspose.Cells for .NET.
// Use Cases: Create a responsive sales chart for a web dashboard by delivering the graphic as SVG. | Include a high‑quality, resolution‑independent chart in documentation or PDF reports. | Provide on‑demand SVG chart generation in a web API that returns vector graphics to clients.
// AI Prompts: Show how to set custom width and height in SvgImageOptions when exporting a chart to SVG. | Provide code to add data labels to the line chart before calling ToImage for SVG output. | Explain how to embed a custom TrueType font in the SVG using SvgEmbeddedFontType.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgChartDemo
{
    // Creates a workbook, fills month‑sales data, adds a line chart, configures SvgImageOptions (fit to viewport, CSS prefix, WOFF font), and renders the chart directly to an SVG file while optionally saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Month");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(210);
                sheet.Cells["B4"].PutValue(150);

                // Add a line chart covering a range of cells
                int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories
                chart.Title.Text = "Quarterly Sales";

                // Configure SVG rendering options
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    // ImageFormat is implicitly SVG for SvgImageOptions, no need to set ImageType
                    FitToViewPort = true,               // Make SVG fit the viewport
                    CssPrefix = "mychart-",             // Optional CSS prefix
                    EmbeddedFontType = SvgEmbeddedFontType.Woff // Embed WOFF font if needed
                };

                // Render the chart directly to an SVG file
                string outputSvgPath = "QuarterlySalesChart.svg";
                chart.ToImage(outputSvgPath, svgOptions);

                // Optionally, save the workbook for reference
                string workbookPath = "QuarterlySalesWorkbook.xlsx";
                workbook.Save(workbookPath);

                Console.WriteLine($"SVG chart generated at: {Path.GetFullPath(outputSvgPath)}");
                Console.WriteLine($"Workbook saved at: {Path.GetFullPath(workbookPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
