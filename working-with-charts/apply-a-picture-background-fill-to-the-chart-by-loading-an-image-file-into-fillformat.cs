using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsChartBackground
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 5, 15, 15);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the chart area fill type to texture (picture)
                chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

                // Load image data from a file if it exists
                string imagePath = "background.png";
                if (File.Exists(imagePath))
                {
                    byte[] imageData = File.ReadAllBytes(imagePath);
                    // Apply the image data to the texture fill of the chart area
                    chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;
                }
                else
                {
                    Console.WriteLine($"Image file not found: {Path.GetFullPath(imagePath)}. Chart will use default background.");
                }

                // Save the workbook
                string outputPath = "ChartWithPictureBackground.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}