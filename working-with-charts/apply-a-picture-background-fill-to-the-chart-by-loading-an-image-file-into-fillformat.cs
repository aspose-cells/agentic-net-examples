using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 5, 15, 15);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the chart area fill type to texture (picture background)
            chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

            string imagePath = "background.png";

            // Load the image file if it exists
            if (File.Exists(imagePath))
            {
                byte[] imageData = File.ReadAllBytes(imagePath);
                chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;
                chart.ChartArea.Area.FillFormat.TextureFill.PictureFormatType = FillPictureType.Stretch;
            }
            else
            {
                Console.WriteLine($"Image file '{imagePath}' not found. Skipping background texture.");
            }

            // Save the workbook with the chart background applied
            workbook.Save("ChartWithPictureBackground.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}