using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsDemo
{
    public class ExportChartToStreamDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // Export the chart to a memory stream and then write the stream to a file
                using (MemoryStream stream = new MemoryStream())
                {
                    // Export chart as PNG image
                    chart.ToImage(stream, ImageType.Png);

                    // Ensure output directory exists
                    string imagePath = "ChartImage.png";
                    string imageDir = Path.GetDirectoryName(imagePath);
                    if (!string.IsNullOrEmpty(imageDir) && !Directory.Exists(imageDir))
                        Directory.CreateDirectory(imageDir);

                    // Write the image bytes to a file
                    File.WriteAllBytes(imagePath, stream.ToArray());
                }

                // Save the workbook (optional)
                string workbookPath = "ChartWorkbook.xlsx";
                string wbDir = Path.GetDirectoryName(workbookPath);
                if (!string.IsNullOrEmpty(wbDir) && !Directory.Exists(wbDir))
                    Directory.CreateDirectory(wbDir);

                workbook.Save(workbookPath);
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}