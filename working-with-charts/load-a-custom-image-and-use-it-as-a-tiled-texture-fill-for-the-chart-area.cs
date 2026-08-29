// Title: How to load a PNG image and apply it as a tiled texture fill to a chart area in Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that reads a PNG file and sets it as a tiled texture fill for the ChartArea of a column chart using Aspose.Cells. | Show how to configure texture scaling, tiling options, and a fallback solid fill when the image file is not found with the Aspose.Cells API. | Demonstrate saving the workbook after applying a custom tiled background to the chart area in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# set chart area background image with tiling | C# Aspose.Cells texture fill for chart area using custom PNG | How to apply tiled texture fill to a chart in Aspose.Cells .NET | Aspose.Cells fallback solid fill when texture image missing | Configure texture scaling and offset for chart background in Aspose.Cells
// Tags: chart background tiled texture Aspose.Cells C# | load png image for chart area Aspose.Cells | configure texture scaling Aspose.Cells chart | fallback solid fill Aspose.Cells when image missing | apply texture fill to chartarea Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace Example
{
    // The example creates a workbook, adds sample data, inserts a column chart, loads a PNG file, and applies it as a tiled texture fill to the chart's background. It configures scaling and tiling, provides a solid‑gray fallback if the image is absent, and saves the file as ChartWithTiledTexture.xlsx.
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

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Apply texture fill if the image file exists
                string texturePath = "customTexture.png";
                if (File.Exists(texturePath))
                {
                    byte[] imageData = File.ReadAllBytes(texturePath);
                    chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;
                    chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;
                    chart.ChartArea.Area.FillFormat.TextureFill.IsTiling = true;
                    chart.ChartArea.Area.FillFormat.TextureFill.TilePicOption = new TilePicOption
                    {
                        ScaleX = 50, // horizontal scale percentage
                        ScaleY = 50, // vertical scale percentage
                        OffsetX = 0,
                        OffsetY = 0
                    };
                }
                else
                {
                    // Fallback to a solid fill when texture image is missing
                    chart.ChartArea.Area.FillFormat.FillType = FillType.Solid;
                    chart.ChartArea.Area.FillFormat.SolidFill.Color = Color.LightGray;
                }

                // Save the workbook with the chart
                workbook.Save("ChartWithTiledTexture.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
