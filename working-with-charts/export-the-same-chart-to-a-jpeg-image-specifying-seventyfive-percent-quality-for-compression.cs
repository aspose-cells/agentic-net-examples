// Title: Export an Aspose.Cells chart to a JPEG file with 75% quality using C#
// Common Searches: Aspose.Cells C# export chart as JPEG with custom quality | How to set JPEG compression level when saving a chart with Aspose.Cells | C# code to export an Excel column chart to JPEG at 75% quality | Using ImageOrPrintOptions to control JPEG quality for chart export in Aspose.Cells | Save Aspose.Cells chart to JPEG image with specific quality setting
// Tags: chart to JPEG ImageOrPrintOptions compression | Aspose.Cells JPEG compression for chart export | export Excel chart as JPEG C# | set JPEG compression Aspose.Cells chart | column chart JPEG export Aspose.Cells | ImageOrPrintOptions JPEG compression setting

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace ExportChartToJpeg
{
    // The example creates a workbook, adds sample data, builds a column chart, configures ImageOrPrintOptions to use JPEG format with a 75% compression quality, and exports the chart to a JPEG file named 'ExportedChart.jpg' using C# and Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
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

                worksheet.Cells["B1"].PutValue("Sales");
                worksheet.Cells["B2"].PutValue(1200);
                worksheet.Cells["B3"].PutValue(800);
                worksheet.Cells["B4"].PutValue(1500);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.SetChartDataRange("A1:B4", true);

                // Configure image export options: JPEG format with 75% quality
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Jpeg,
                    Quality = 75 // Quality range: 0 (lowest) to 100 (highest)
                };

                // Export the chart to a JPEG file using the specified options
                string outputPath = "ExportedChart.jpg";
                chart.ToImage(outputPath, imgOptions);

                Console.WriteLine($"Chart successfully exported to '{outputPath}' with 75% JPEG quality.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while exporting the chart:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
