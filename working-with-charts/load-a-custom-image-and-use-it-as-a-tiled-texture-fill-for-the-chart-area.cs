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
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set the chart area fill type to texture
            chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

            // Load custom image bytes if the file exists
            string texturePath = "customTexture.png";
            if (File.Exists(texturePath))
            {
                try
                {
                    byte[] imageData = File.ReadAllBytes(texturePath);
                    chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;
                    chart.ChartArea.Area.FillFormat.TextureFill.IsTiling = true;

                    // Configure tile scaling and offset
                    chart.ChartArea.Area.FillFormat.TextureFill.TilePicOption = new TilePicOption
                    {
                        ScaleX = 50, // 50% horizontal scaling
                        ScaleY = 50, // 50% vertical scaling
                        OffsetX = 0,
                        OffsetY = 0
                    };
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to load texture: {ex.Message}");
                    // Fallback to solid fill if texture loading fails
                    chart.ChartArea.Area.FillFormat.FillType = FillType.Solid;
                }
            }
            else
            {
                // If texture file not found, use solid fill
                chart.ChartArea.Area.FillFormat.FillType = FillType.Solid;
            }

            // Save the workbook
            string outputPath = "ChartWithTexture.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}