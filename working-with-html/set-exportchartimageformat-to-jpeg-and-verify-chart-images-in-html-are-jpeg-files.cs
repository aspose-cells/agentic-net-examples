// Title: Export an Excel chart to JPEG and embed it in an HTML page using Aspose.Cells for .NET
// Common Searches: how to export an Excel chart as a JPEG using Aspose.Cells in C# | create HTML page that displays a chart image saved by Aspose.Cells | check if a saved chart image is a JPEG file in .NET | set chart image format to JPEG when exporting with Aspose.Cells | Aspose.Cells chart ToImage method JPEG example
// Tags: Aspose.Cells chart export JPEG | C# embed chart image in HTML | chart.ToImage ImageType.Jpeg usage | validate JPEG header .NET | generate HTML with Aspose.Cells chart image

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace ExportChartImageFormatDemo
{
    // The sample creates a workbook, adds a column chart, exports the chart to a JPEG file with chart.ToImage, generates a simple HTML page that references the JPEG, and verifies the saved file is a valid JPEG by checking its header.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a workbook with sample data and a chart
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["A3"].PutValue("Orange");
                sheet.Cells["A4"].PutValue("Banana");

                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(1200);
                sheet.Cells["B3"].PutValue(800);
                sheet.Cells["B4"].PutValue(1500);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";
                chart.Title.Text = "Fruit Sales";

                // 2. Export the chart as a JPEG image
                string imageFolder = "ChartImages";
                Directory.CreateDirectory(imageFolder);
                string chartImagePath = Path.Combine(imageFolder, "ChartImage.jpg");

                // Save chart as JPEG
                chart.ToImage(chartImagePath, ImageType.Jpeg);
                Console.WriteLine($"Chart image saved as JPEG: {chartImagePath}");

                // 3. Generate simple HTML that references the JPEG image
                string htmlFile = "WorkbookWithChart.html";
                string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset=""utf-8"" />
    <title>Workbook with Chart</title>
</head>
<body>
    <h2>{chart.Title.Text}</h2>
    <img src=""{Path.Combine(imageFolder, "ChartImage.jpg")}"" alt=""Chart Image"" />
</body>
</html>";
                File.WriteAllText(htmlFile, htmlContent);
                Console.WriteLine($"HTML file created: {htmlFile}");

                // 4. Verify that the exported chart image is a JPEG file
                if (!File.Exists(chartImagePath))
                {
                    Console.WriteLine($"Chart image not found: {chartImagePath}");
                    return;
                }

                bool isJpeg = IsJpegFile(chartImagePath);
                string ext = Path.GetExtension(chartImagePath).ToLowerInvariant();
                Console.WriteLine($"Verification -> File: {Path.GetFileName(chartImagePath)} | Extension: {ext} | JPEG Header: {(isJpeg ? "OK" : "FAIL")}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        /// <param name="filePath">Full path to the image file.</param>
        /// <returns>True if the file is a JPEG; otherwise false.</returns>
        private static bool IsJpegFile(string filePath)
        {
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    if (fs.Length < 2)
                        return false;

                    int firstByte = fs.ReadByte();
                    int secondByte = fs.ReadByte();

                    // JPEG files start with 0xFF 0xD8
                    return firstByte == 0xFF && secondByte == 0xD8;
                }
            }
            catch
            {
                return false;
            }
        }
    }
}
