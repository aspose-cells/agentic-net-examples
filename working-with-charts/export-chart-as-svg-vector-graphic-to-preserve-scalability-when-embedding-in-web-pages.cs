// Title: Export an Aspose.Cells line chart to an SVG file using C# for web‑ready vector graphics
// AI Prompts: Generate C# code that creates a line chart in an Aspose.Cells workbook and saves it as an SVG using ImageOrPrintOptions. | Show how to export any Aspose.Cells chart to SVG with a transparent background and custom image dimensions. | Adapt the example to export a column chart to SVG and set the output file path dynamically.
// Common Searches: aspnet export excel chart as svg with aspose.cells | c# save aspose.cells chart to scalable vector graphic file | how to use ImageOrPrintOptions to generate svg from aspose chart | convert line chart from workbook to svg for responsive web page | aspose.cells chart toimage svg output example c#
// Tags: chart.ToImage SVG export Aspose.Cells | Aspose.Cells line chart SVG output | C# export Excel chart as vector graphic | ImageOrPrintOptions SVG configuration Aspose | scalable chart rendering Aspose.Cells C#

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgExport
{
    // The sample creates a workbook, fills month and sales data, adds a line chart, and uses ImageOrPrintOptions with chart.ToImage to write the chart to 'QuarterlySalesChart.svg', producing a scalable SVG suitable for web embedding.
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

                // Add a line chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories
                chart.Title.Text = "Quarterly Sales";

                // Export the chart as an SVG file
                string outputPath = "QuarterlySalesChart.svg";

                // ImageOrPrintOptions without explicit ImageFormat; format is inferred from file extension
                ImageOrPrintOptions options = new ImageOrPrintOptions();

                chart.ToImage(outputPath, options);

                Console.WriteLine($"Chart exported successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
