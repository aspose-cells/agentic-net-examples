// Title: Apply Wood Grain Texture Fill to a Chart Area with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a column chart, and fills the chart area with a tiled wood‑grain PNG texture using Aspose.Cells. If the image is missing, a solid light‑gray fill is applied, and the workbook is saved as ChartWithWoodTexture.xlsx.
// Keywords: Aspose.Cells chart texture fill | C# chart area background image | FillType.Texture example | Excel chart wood grain texture | Aspose.Cells FillFormat texture | chart area tiled image .NET
// Common Searches: Aspose.Cells set chart area texture image | C# apply PNG texture to Excel chart background | how to use FillType.Texture in Aspose.Cells | chart area background wood grain Aspose.Cells | fallback solid fill for missing chart texture
// Developer Intent: Set a wood‑grain image as a tiled texture fill for a chart’s background, with a solid‑color fallback when the image is unavailable.
// Use Cases: Generate a column chart and give it a natural wood‑grain appearance. | Provide a reliable fallback color if the texture file cannot be found. | Export the styled chart to an Excel file for distribution or further processing.
// AI Prompts: Show C# code that loads a PNG file and applies it as a tiled texture fill to a chart area using Aspose.Cells. | Generate a snippet that checks for a texture image, applies it, and falls back to a solid color when missing. | Explain how to adjust tiling, scaling, and rotation of a texture fill on an Aspose.Cells chart.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing; // Needed for FillType enum

namespace AsposeCellsTextureFillDemo
{
    // Creates a workbook, adds a column chart, and fills the chart area with a tiled wood‑grain PNG texture using Aspose.Cells. If the image is missing, a solid light‑gray fill is applied, and the workbook is saved as ChartWithWoodTexture.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
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

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Apply a texture fill to the chart area using a wood grain image
                chart.ChartArea.Area.FillFormat.FillType = FillType.Texture;

                // Load the wood grain image file if it exists
                string imagePath = Path.Combine(Environment.CurrentDirectory, "wood_grain.png");
                if (File.Exists(imagePath))
                {
                    byte[] imageData = File.ReadAllBytes(imagePath);
                    chart.ChartArea.Area.FillFormat.TextureFill.ImageData = imageData;
                    chart.ChartArea.Area.FillFormat.TextureFill.IsTiling = true; // repeat texture
                }
                else
                {
                    // Fallback to a solid fill if the image is missing
                    chart.ChartArea.Area.FillFormat.FillType = FillType.Solid;
                    chart.ChartArea.Area.FillFormat.SolidFill.Color = System.Drawing.Color.LightGray;
                }

                // Save the workbook
                string outputPath = Path.Combine(Environment.CurrentDirectory, "ChartWithWoodTexture.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
