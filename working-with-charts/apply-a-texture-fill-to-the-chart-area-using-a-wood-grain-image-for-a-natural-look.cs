using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureFillDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("North");
            sheet.Cells["A3"].PutValue("South");
            sheet.Cells["A4"].PutValue("East");
            sheet.Cells["A5"].PutValue("West");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["B3"].PutValue(850);
            sheet.Cells["B4"].PutValue(950);
            sheet.Cells["B5"].PutValue(1100);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Apply a texture fill to the chart area using a wood grain image
            // Ensure the chart area uses texture fill type
            chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

            // Load the wood grain image bytes (replace with actual file path)
            string imagePath = Path.Combine(Environment.CurrentDirectory, "wood_grain.png");
            if (File.Exists(imagePath))
            {
                byte[] imageData = File.ReadAllBytes(imagePath);
                chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;
                // Optional: tile the texture for a seamless look
                chart.ChartArea.Area.FillFormat.TextureFill.IsTiling = true;
            }
            else
            {
                Console.WriteLine($"Image file not found: {imagePath}");
            }

            // Save the workbook
            workbook.Save("ChartWithWoodTexture.xlsx");
        }
    }
}