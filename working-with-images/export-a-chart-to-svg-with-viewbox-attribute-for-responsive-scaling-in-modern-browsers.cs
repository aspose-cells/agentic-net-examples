// Title: Export an Excel column chart to responsive SVG with a viewBox using Aspose.Cells for .NET (C#)
// Common Searches: how to generate a responsive SVG chart with viewBox from Aspose.Cells in C# | Aspose.Cells C# export Excel chart to SVG with viewBox attribute | C# code to render Excel column chart as SVG that scales with viewBox | using ImageOrPrintOptions to create SVG chart with viewBox in Aspose.Cells | export chart to SVG for web using Aspose.Cells .NET library
// Tags: Aspose.Cells chart SVG export with viewBox | ImageOrPrintOptions SVGFitToViewPort setting | C# responsive SVG chart generation from Excel | column chart to SVG using Aspose.Cells | SVG chart scaling via viewBox in .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace ExportChartToSvgWithViewBox
{
    // The example creates a workbook, adds sample fruit‑sales data, builds a column chart, and configures ImageOrPrintOptions with SVGFitToViewPort to embed a viewBox attribute. The chart is rendered to a responsive SVG file (ChartOutput.svg) and the code verifies the presence of the viewBox, demonstrating how to export Excel charts as scalable SVG graphics using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a new workbook and populate it with sample data.
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Header
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // Sample rows
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(45);
                sheet.Cells["A4"].PutValue("Cherry");
                sheet.Cells["B4"].PutValue(25);

                // -----------------------------------------------------------------
                // 2. Add a column chart that uses the data above.
                // -----------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories
                chart.Title.Text = "Fruit Sales";

                // -----------------------------------------------------------------
                // 3. Configure SVG rendering options.
                // -----------------------------------------------------------------
                ImageOrPrintOptions svgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Svg,
                    OnePagePerSheet = true,
                    SVGFitToViewPort = true,   // Generates viewBox (obsolete but functional)
                    SvgCssPrefix = "chart_"    // Optional CSS prefix (obsolete but functional)
                };

                // -----------------------------------------------------------------
                // 4. Render the chart to an SVG file.
                // -----------------------------------------------------------------
                string svgPath = "ChartOutput.svg";
                chart.ToImage(svgPath, svgOptions);
                Console.WriteLine($"Chart exported to SVG: {svgPath}");

                // -----------------------------------------------------------------
                // 5. Verify that the SVG contains a viewBox attribute.
                // -----------------------------------------------------------------
                if (File.Exists(svgPath))
                {
                    string svgContent = File.ReadAllText(svgPath);
                    bool hasViewBox = svgContent.Contains("viewBox");
                    Console.WriteLine($"SVG contains viewBox attribute: {hasViewBox}");

                    // Show a snippet of the SVG header for confirmation.
                    int headerEnd = svgContent.IndexOf('>');
                    if (headerEnd > 0)
                    {
                        Console.WriteLine("SVG header snippet:");
                        Console.WriteLine(svgContent.Substring(0, Math.Min(200, headerEnd + 1)));
                    }
                }
                else
                {
                    Console.WriteLine($"Error: SVG file not found at path '{svgPath}'.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            // -----------------------------------------------------------------
            // 6. Keep console window open when run from IDE.
            // -----------------------------------------------------------------
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
