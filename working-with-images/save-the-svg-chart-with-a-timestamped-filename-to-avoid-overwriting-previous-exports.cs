// Title: Save an Excel chart as an SVG file with a timestamped name using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export chart to SVG with unique filename | How to add a timestamp to the SVG file name when saving an Excel chart in .NET | Aspose.Cells generate SVG chart file without overwriting previous exports | Create timestamped SVG output for an Excel chart using ImageOrPrintOptions in C# | Save Excel chart as SVG and avoid filename collisions with Aspose.Cells
// Tags: chart to SVG export Aspose.Cells | unique file naming C# | prevent file name collisions Aspose.Cells | ImageOrPrintOptions SVG configuration | programmatic Excel chart export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgExport
{
    // The example loads an existing workbook or creates a new one with sample data and a column chart, configures ImageOrPrintOptions for SVG output, builds a filename that includes the current date and time (ChartExport_yyyyMMdd_HHmmssfff.svg) to guarantee uniqueness, and then exports the first chart to an SVG file while handling potential errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to an existing workbook (optional). If the file does not exist, a new workbook will be created.
                string inputPath = "input.xlsx";

                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    // Create a new workbook with sample data and a chart
                    workbook = new Workbook();
                    Worksheet sheet = workbook.Worksheets[0];

                    // Populate sample data
                    sheet.Cells["A1"].PutValue("Category");
                    sheet.Cells["B1"].PutValue("Value");
                    sheet.Cells["A2"].PutValue("A");
                    sheet.Cells["B2"].PutValue(10);
                    sheet.Cells["A3"].PutValue("B");
                    sheet.Cells["B3"].PutValue(20);
                    sheet.Cells["A4"].PutValue("C");
                    sheet.Cells["B4"].PutValue(30);

                    // Add a column chart
                    int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                    Chart chart = sheet.Charts[chartIndex];
                    chart.NSeries.Add("B2:B4", true);
                    chart.NSeries.CategoryData = "A2:A4";
                    chart.Title.Text = "Sample Column Chart";

                    Console.WriteLine("Created new workbook with sample data and chart.");
                }

                // Retrieve the first chart from the first worksheet
                Worksheet firstSheet = workbook.Worksheets[0];
                if (firstSheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the workbook. Exiting.");
                    return;
                }

                Chart targetChart = firstSheet.Charts[0];

                // Configure image options for SVG output
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Svg,
                    OnePagePerSheet = true,
                    // Optional: make the SVG fit the viewport (obsolete but still functional)
                    SVGFitToViewPort = true
                };

                // Generate a timestamped filename to avoid overwriting previous exports
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmssfff");
                string outputFileName = $"ChartExport_{timestamp}.svg";

                // Export the chart to SVG
                try
                {
                    targetChart.ToImage(outputFileName, imgOptions);
                    Console.WriteLine($"Chart successfully exported to SVG file: {outputFileName}");
                }
                catch (Exception exportEx)
                {
                    Console.WriteLine("Failed to export chart to SVG:");
                    Console.WriteLine(exportEx.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during SVG export:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
