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

            // Sample data for the chart
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

            // Set chart area fill to texture
            chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

            // Load custom texture if the file exists
            string imagePath = "customTexture.png";
            if (File.Exists(imagePath))
            {
                byte[] imageData = File.ReadAllBytes(imagePath);
                chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;
                chart.ChartArea.Area.FillFormat.TextureFill.IsTiling = true;

                // Configure tile scaling (percentage of original size)
                TilePicOption tileOption = new TilePicOption
                {
                    ScaleX = 50, // 50% horizontal scale
                    ScaleY = 50  // 50% vertical scale
                };
                chart.ChartArea.Area.FillFormat.TextureFill.TilePicOption = tileOption;
            }
            else
            {
                // Fallback: use a solid fill if texture not found
                chart.ChartArea.Area.FillFormat.FillType = FillType.Solid;
                chart.ChartArea.Area.FillFormat.SolidFill.Color = System.Drawing.Color.LightGray;
            }

            // Save the workbook
            string outputPath = "ChartWithTiledTexture.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}