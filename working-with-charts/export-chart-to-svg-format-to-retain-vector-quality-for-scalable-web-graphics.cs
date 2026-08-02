// Title: Export Aspose.Cells Chart to Scalable SVG in C# Using SvgImageOptions
// Description: Shows how to build a workbook, add a line chart, set SvgImageOptions (FitToViewPort, CssPrefix, EmbeddedFontType), and write the chart directly to an SVG file, with optional workbook saving for further processing.
// Keywords: Aspose.Cells | C# SVG chart export | SvgImageOptions | chart.ToImage | vector SVG | line chart export | Excel to SVG | Aspose.Cells chart rendering | scalable vector graphics | .NET chart SVG
// Common Searches: Aspose.Cells export chart as SVG C# | How to use SvgImageOptions with Aspose.Cells | C# generate SVG from Excel chart | Save Aspose.Cells chart to SVG file | Vector chart export Aspose.Cells .NET
// Developer Intent: Create an SVG representation of an Aspose.Cells chart in C# while preserving vector quality.
// Use Cases: Render a high‑resolution line chart for a responsive web dashboard. | Produce printable SVG graphics for reports that require crisp scaling. | Extract chart visuals as SVG assets for reuse in other design tools while keeping the original workbook editable.
// AI Prompts: Write C# code that builds a bar chart with Aspose.Cells and exports it to SVG using custom CSS classes and embedded fonts. | Explain how to loop through all charts in a workbook and save each as a separate SVG file with distinct SvgImageOptions settings. | Describe the benefits of embedding WOFF fonts in SVG output from Aspose.Cells and show the required option configuration.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgExport
{
    // Shows how to build a workbook, add a line chart, set SvgImageOptions (FitToViewPort, CssPrefix, EmbeddedFontType), and write the chart directly to an SVG file, with optional workbook saving for further processing.
    public class ExportChartToSvg
    {
        public static void Run()
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

                // Add a line chart covering the data range
                int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories
                chart.Title.Text = "Quarterly Sales";

                // Configure SVG rendering options
                SvgImageOptions svgOptions = new SvgImageOptions
                {
                    FitToViewPort = true,                     // Make SVG fit the viewport (optional)
                    CssPrefix = "chart-",                     // Optional CSS prefix for styling
                    EmbeddedFontType = SvgEmbeddedFontType.Woff // Optional font embedding
                };

                // Export the chart directly to an SVG file
                string outputPath = "QuarterlySalesChart.svg";
                chart.ToImage(outputPath, svgOptions);

                // Optionally save the workbook for reference
                string workbookPath = "QuarterlySalesWorkbook.xlsx";
                workbook.Save(workbookPath);

                Console.WriteLine($"Chart exported to SVG: {outputPath}");
                Console.WriteLine($"Workbook saved to: {workbookPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportChartToSvg.Run();
        }
    }
}
