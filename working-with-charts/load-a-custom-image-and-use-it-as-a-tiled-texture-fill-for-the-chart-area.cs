// Title: Aspose.Cells for .NET – Apply a Custom Tiled Texture Fill to a Chart Area (C#)
// Description: Demonstrates how to create a workbook, add a column chart, load a PNG image as a byte array, set the chart area's FillFormat to Texture, enable tiling, and fine‑tune tile scaling with TilePicOption before saving the Excel file.
// Keywords: Aspose.Cells chart texture fill | C# tiled background image chart | TilePicOption scaling Aspose.Cells | load image bytes chart fill .NET | Excel chart area texture fill | custom chart background image | Aspose.Cells chart area fill format
// Common Searches: Aspose.Cells custom image tiled texture for chart area | How to set TilePicOption scale on chart background in C# | Load PNG into chart fill format Aspose.Cells | Enable texture tiling for Excel chart using Aspose | C# example of chart area texture fill with Aspose.Cells
// Developer Intent: Create a column chart and use a user‑provided PNG as a repeating texture for the chart’s background.
// Use Cases: Apply a branded pattern or watermark as a tiled background behind chart data. | Adjust tile size and proportion by configuring ScaleX and ScaleY on TilePicOption. | Gracefully handle missing image files while still generating a valid workbook.
// AI Prompts: Write C# code with Aspose.Cells that loads a PNG file and applies it as a tiled texture fill to a chart area, including optional scaling parameters. | Explain how TilePicOption properties (ScaleX, ScaleY, OffsetX, OffsetY) affect the appearance of a texture fill in an Aspose.Cells chart. | Suggest best‑practice error handling for missing texture images before saving an Excel workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureFillDemo
{
    // Demonstrates how to create a workbook, add a column chart, load a PNG image as a byte array, set the chart area's FillFormat to Texture, enable tiling, and fine‑tune tile scaling with TilePicOption before saving the Excel file.
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
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = worksheet.Charts[chartIndex];

                // Populate sample data for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(100);
                worksheet.Cells["B3"].PutValue(200);
                worksheet.Cells["B4"].PutValue(300);
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Set the fill type of the chart area to texture
                chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

                // Load custom image data if the file exists
                string imagePath = Path.Combine(Environment.CurrentDirectory, "customTexture.png");
                if (File.Exists(imagePath))
                {
                    byte[] imageData = File.ReadAllBytes(imagePath);
                    chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;
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
                }
                else
                {
                    Console.WriteLine($"Image file not found: {imagePath}. Skipping texture fill.");
                }

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
