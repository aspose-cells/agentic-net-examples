// Title: Apply a Custom Tiled Texture Fill to an Excel Chart Area with Aspose.Cells for .NET
// Description: Shows how to read a PNG (or a base‑64 placeholder), set the chart area's FillFormat to Texture, enable tiling, and fine‑tune scaling and offset via TilePicOption, then save the workbook.
// Keywords: Aspose.Cells chart texture fill | tiled background image Excel chart .NET | load custom PNG Aspose.Cells | TextureFill TilePicOption scaling | chart area FillFormat Texture Aspose
// Common Searches: Aspose.Cells set tiled texture for chart area | how to use a PNG as chart background in .NET | enable texture tiling on Excel chart with Aspose | adjust tile scaling offset in Aspose.Cells chart | fallback image for missing chart texture Aspose
// Developer Intent: Load an image and apply it as a repeating texture fill to the chart area of an Excel chart using Aspose.Cells.
// Use Cases: Create a column chart and give it a patterned background using a user‑provided PNG. | Provide a default red‑pixel image when the specified texture file cannot be found. | Control the size and position of each tile with TilePicOption properties.
// AI Prompts: Write C# code that reads a user‑specified image file and applies it as a tiled texture fill to an Aspose.Cells chart area, including error handling for missing files. | Explain how to modify TilePicOption.ScaleX, ScaleY, OffsetX, and OffsetY to customize the appearance of a tiled chart background. | Show how to replace the base‑64 placeholder with another image and update the texture fill without rebuilding the workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureFillDemo
{
    // Shows how to read a PNG (or a base‑64 placeholder), set the chart area's FillFormat to Texture, enable tiling, and fine‑tune scaling and offset via TilePicOption, then save the workbook.
    class Program
    {
        static void Main()
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

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure chart area to use texture fill
                chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

                // Load custom texture image or use a placeholder
                string imagePath = Path.Combine(Environment.CurrentDirectory, "customTexture.png");
                byte[] imageData;

                if (File.Exists(imagePath))
                {
                    // Load image bytes from file
                    imageData = File.ReadAllBytes(imagePath);
                }
                else
                {
                    // 1x1 red PNG (base64 encoded) as a placeholder
                    const string redPixelBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
                    imageData = Convert.FromBase64String(redPixelBase64);
                }

                // Apply the image as the texture fill
                chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;

                // Enable tiling so the image repeats across the chart area
                chart.ChartArea.Area.FillFormat.TextureFill.IsTiling = true;

                // Optional: adjust tile scaling and offset
                TilePicOption tileOptions = new TilePicOption
                {
                    ScaleX = 50,   // 50% horizontal scale
                    ScaleY = 50,   // 50% vertical scale
                    OffsetX = 0,
                    OffsetY = 0
                };
                chart.ChartArea.Area.FillFormat.TextureFill.TilePicOption = tileOptions;

                // Save the workbook
                string outputPath = Path.Combine(Environment.CurrentDirectory, "ChartWithTiledTexture.xlsx");
                workbook.Save(outputPath);

                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
