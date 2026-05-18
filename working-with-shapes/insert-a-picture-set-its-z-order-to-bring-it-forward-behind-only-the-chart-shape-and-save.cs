using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsExample
{
    class InsertPictureAndAdjustZOrder
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 1, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Load image data if the file exists
                const string imagePath = "example.jpg";
                if (File.Exists(imagePath))
                {
                    byte[] imageBytes = File.ReadAllBytes(imagePath);
                    using (MemoryStream imageStream = new MemoryStream(imageBytes))
                    {
                        // Insert picture into the chart (position and scale are in 1/4000 of chart area)
                        Picture picture = chart.Shapes.AddPictureInChart(
                            top: 100,          // vertical offset
                            left: 100,         // horizontal offset
                            stream: imageStream,
                            widthScale: 50,    // 50% of original width
                            heightScale: 50);  // 50% of original height

                        // Bring the picture forward one position in Z-order.
                        picture.ToFrontOrBack(1);
                    }
                }
                else
                {
                    Console.WriteLine($"Image file '{imagePath}' not found. Skipping picture insertion.");
                }

                // Save the workbook
                workbook.Save("InsertPictureZOrderDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}