// Title: Apply Wood Grain Texture Fill to a Chart Area with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, sets the chart area's FillFormat to Texture, loads a wood_grain.png image, applies it as a tiled background, and saves the file as ChartWithWoodTexture.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart texture fill | FillType.Texture | chart area background image | wood grain texture | .NET Excel chart styling | image texture fill Aspose.Cells | Excel chart background image | Aspose.Cells chart area fill
// Common Searches: Aspose.Cells set chart area texture fill C# | how to use PNG as chart background in Aspose.Cells | apply wood grain image to Excel chart area with Aspose | enable tiling for chart texture fill Aspose.Cells | chart area FillFormat.Texture example .NET
// Developer Intent: Add a custom wood‑grain image as the textured background of a chart area.
// Use Cases: Design printable sales reports with a wooden‑styled chart background. | Maintain brand consistency by applying product‑specific textures to Excel charts. | Create themed dashboards where chart areas repeat a texture pattern.
// AI Prompts: Generate C# code that loads an image from a file stream and sets it as a tiled texture fill for a chart area in Aspose.Cells. | Show how to apply a JPEG texture to a chart area and disable tiling using Aspose.Cells for .NET. | Explain best practices for handling missing image files when assigning a texture fill to a chart.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureFillDemo
{
    // Creates a workbook, adds sample data, inserts a column chart, sets the chart area's FillFormat to Texture, loads a wood_grain.png image, applies it as a tiled background, and saves the file as ChartWithWoodTexture.xlsx using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Wood");
            sheet.Cells["A3"].PutValue("Metal");
            sheet.Cells["A4"].PutValue("Plastic");

            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["B4"].PutValue(150);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply a texture fill to the chart area using a wood grain image
            // Set the fill type to Texture
            chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

            // Load the wood grain image into a byte array (replace with your actual image path)
            string imagePath = Path.Combine(Environment.CurrentDirectory, "wood_grain.png");
            if (File.Exists(imagePath))
            {
                byte[] imageData = File.ReadAllBytes(imagePath);
                // Assign the image data to the texture fill
                chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;

                // Optional: enable tiling if you want the texture to repeat
                chart.ChartArea.Area.FillFormat.TextureFill.IsTiling = true;
            }
            else
            {
                Console.WriteLine("Image file not found: " + imagePath);
            }

            // Save the workbook
            workbook.Save("ChartWithWoodTexture.xlsx");
        }
    }
}
