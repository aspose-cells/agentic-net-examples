// Title: Export a column chart to JPEG with specific quality and validate the file size range using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export chart to JPEG with quality setting | How to verify JPEG file size after chart export in .NET | Validate image size range for exported chart using Aspose.Cells | C# code to ensure chart JPEG is between 10KB and 500KB | Set JPEG quality when saving an Aspose.Cells chart and check output size
// Tags: chart export jpeg quality Aspose.Cells | jpeg file size validation .NET | image export options Aspose.Cells C# | column chart to jpeg Aspose.Cells | file size range check C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsJpegQualityValidator
{
    // The example creates a workbook, adds sample data, builds a column chart, and exports the chart to a JPEG file using ImageOrPrintOptions with a quality of 80. After saving, it confirms the file exists, reads its size, and verifies that the size falls between 10 KB and 500 KB, reporting success or failure.
    class Program
    {
        // Define acceptable file size range in bytes (example: 10 KB to 500 KB)
        private const long MinFileSizeBytes = 10 * 1024;   // 10 KB
        private const long MaxFileSizeBytes = 500 * 1024; // 500 KB

        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and populate it with sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["B4"].PutValue(150);

                // 2. Add a column chart based on the data
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // 3. Prepare image export options for JPEG with a specific quality
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    ImageType = ImageType.Jpeg,
                    Quality = 80 // JPEG quality (0-100)
                };

                // 4. Define output file path
                string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "ExportedChart.jpg");

                // 5. Export the chart to JPEG using the options
                chart.ToImage(outputFile, imgOptions);

                // 6. Verify that the file was created
                if (!File.Exists(outputFile))
                {
                    Console.WriteLine($"Error: The image file was not created at '{outputFile}'.");
                    return;
                }

                // 7. Get the file size
                FileInfo fileInfo = new FileInfo(outputFile);
                long fileSize = fileInfo.Length;

                // 8. Validate file size against the acceptable range
                if (fileSize >= MinFileSizeBytes && fileSize <= MaxFileSizeBytes)
                {
                    Console.WriteLine($"Success: JPEG image size ({fileSize} bytes) is within the acceptable range [{MinFileSizeBytes}, {MaxFileSizeBytes}] bytes.");
                }
                else
                {
                    Console.WriteLine($"Failure: JPEG image size ({fileSize} bytes) is outside the acceptable range [{MinFileSizeBytes}, {MaxFileSizeBytes}] bytes.");
                }

                // Optional: Output the path for manual inspection
                Console.WriteLine($"Image saved at: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An exception occurred: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
