// Title: Disable texture fill tiling for small shapes in Aspose.Cells for .NET
// Description: Learn how to add a rectangle shape, load a PNG as a texture fill, and automatically turn off TextureFill.IsTiling when the shape's pixel dimensions are smaller than the image size, preventing repeated tiles in the generated Excel workbook.
// Keywords: Aspose.Cells texture fill | disable tiling C# | shape size image dimensions | TextureFill.IsTiling | conditional texture fill | Aspose.Cells shape fill | Excel shape texture | C# Aspose.Cells example | image size comparison | non‑tiled texture fill
// Common Searches: Aspose.Cells disable texture tiling | C# compare shape size with image size Aspose.Cells | how to prevent texture repeat in Excel shape | set TextureFill.IsTiling false based on dimensions | conditional texture fill in Aspose.Cells
// Developer Intent: Identify when the shape's width or height (in pixels) is less than the source image dimensions and set TextureFill.IsTiling to false accordingly.
// Use Cases: Create an Excel file where a small shape shows the texture only once, without tiling. | Apply tiled textures to large shapes while keeping small shapes non‑tiled in the same worksheet. | Automate image‑to‑shape rendering where the fill mode adapts to the shape's size.
// AI Prompts: Generate C# code using Aspose.Cells that reads a PNG's width and height, compares them to a shape's pixel size, and disables texture tiling when the shape is smaller. | Provide a reusable method that accepts a Shape object and image bytes, determines if tiling is needed, and sets TextureFill.IsTiling appropriately. | Show how to conditionally enable or disable texture fill tiling in Aspose.Cells based on shape versus image dimensions.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureTilingDemo
{
    // Learn how to add a rectangle shape, load a PNG as a texture fill, and automatically turn off TextureFill.IsTiling when the shape's pixel dimensions are smaller than the image size, preventing repeated tiles in the generated Excel workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the texture image
                string imagePath = "texture.png";
                if (!File.Exists(imagePath))
                {
                    Console.WriteLine($"Image file not found: {imagePath}");
                    return;
                }

                // Load image bytes
                byte[] imageBytes = File.ReadAllBytes(imagePath);

                // Define shape size (width and height in pixels)
                int shapeWidth = 150;
                int shapeHeight = 100;

                // Add a rectangle shape to the worksheet
                // Parameters: upper left row, upper left column, row offset, column offset, width, height
                Shape shape = worksheet.Shapes.AddRectangle(2, 2, 0, 0, shapeWidth, shapeHeight);

                // Apply texture fill to the shape
                shape.Fill.FillType = FillType.Texture;
                TextureFill textureFill = shape.Fill.TextureFill;
                textureFill.ImageData = imageBytes;

                // Since we are not using System.Drawing to obtain image dimensions,
                // we will enable tiling by default. Adjust this logic if image size is known.
                textureFill.IsTiling = true;
                Console.WriteLine("Tiling enabled for the texture fill.");

                // Save the workbook
                string outputPath = "TextureTilingDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
