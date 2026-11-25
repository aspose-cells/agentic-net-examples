// Title: Export a column chart to an SVG vector graphic with Aspose.Cells for .NET (C#)
// Common Searches: how to save an Excel column chart as SVG using Aspose.Cells C# | Aspose.Cells C# export chart to scalable SVG image | C# code example for chart.ToImage SVG output Aspose.Cells | set custom CSS prefix when exporting chart to SVG with Aspose.Cells | use ImageOrPrintOptions to generate SVG from Excel chart in .NET
// Tags: chart to SVG conversion Aspose.Cells | ImageOrPrintOptions SVG settings Aspose.Cells | C# export Excel chart as SVG file | custom SvgCssPrefix configuration Aspose.Cells | OnePagePerSheet rendering option for chart SVG | column chart SVG output .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing; // For ImageType

namespace ExportChartToSvgDemo
{
    // The sample creates a workbook, adds a column chart, configures ImageOrPrintOptions for SVG output (including a custom CSS prefix and OnePagePerSheet), and uses chart.ToImage to export the chart as a scalable SVG file, then verifies the file creation.
    class Program
    {
        static void Main(string[] args)
        {
            // Define output paths
            string outputDirectory = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            string svgFilePath = Path.Combine(outputDirectory, "SampleChart.svg");

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDirectory);

            try
            {
                // 1. Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["A4"].PutValue("Banana");
                sheet.Cells["B4"].PutValue(150);

                // 2. Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.CategoryData = "A2:A4";
                chart.NSeries.Add("B2:B4", true);
                chart.Title.Text = "Fruit Sales";

                // 3. Configure image options for SVG output
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Svg,      // Specify SVG output
                    OnePagePerSheet = true          // Render whole sheet/chart on one page
                };

                // The newer SvgImageOptions property may not be available in older versions.
                // Use the legacy (still functional) properties instead.
                imgOptions.SvgCssPrefix = "myChart_";
                imgOptions.SVGFitToViewPort = true;

                // 4. Export the chart to an SVG file
                chart.ToImage(svgFilePath, imgOptions);
                Console.WriteLine($"Chart successfully exported to SVG at: {svgFilePath}");

                // 5. Verify the SVG file was created and contains the CSS prefix
                if (File.Exists(svgFilePath))
                {
                    string svgContent = File.ReadAllText(svgFilePath);
                    bool containsPrefix = svgContent.Contains("myChart_");
                    Console.WriteLine($"SVG file contains CSS prefix \"myChart_\": {containsPrefix}");
                }
                else
                {
                    Console.WriteLine("Error: SVG file was not created.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while exporting the chart to SVG:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
