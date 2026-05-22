using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class AddPictureToChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);

            // Add a column chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 1, 20, 10);
            Chart chart = sheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B3", true);
            chart.NSeries.CategoryData = "A2:A3";

            // Path to the image file
            string imagePath = "example.jpg";

            // Load image data if the file exists
            if (File.Exists(imagePath))
            {
                byte[] imgBytes = File.ReadAllBytes(imagePath);
                using (MemoryStream imgStream = new MemoryStream(imgBytes))
                {
                    // Add picture to the chart; 100% scaling keeps original size
                    Picture pic = chart.Shapes.AddPictureInChart(100, 100, imgStream, 100, 100);
                    pic.WidthPt = 100; // Set picture width in points
                }
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}. Skipping picture insertion.");
            }

            // Save the workbook
            workbook.Save("ChartWithPicture.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}