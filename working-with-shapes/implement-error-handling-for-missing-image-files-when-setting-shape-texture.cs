// Title: C# – Handle Missing Image File When Applying Texture Fill to a Shape in Aspose.Cells
// Description: Demonstrates how to add a rectangle shape to a workbook, load a PNG as a texture fill, verify the file exists, fall back to the built‑in WaterDroplets texture if it doesn't, and save the file with comprehensive error handling.
// Keywords: Aspose.Cells texture fill | C# shape texture error handling | FileNotFoundException Aspose.Cells | built‑in texture fallback | WaterDroplets texture type | shape fill scaling and tiling | workbook save exception handling | Aspose.Cells .NET example
// Common Searches: Aspose.Cells set shape texture from file | C# handle missing image for shape texture fill | fallback to built‑in texture Aspose.Cells | texture fill FileNotFoundException handling | how to apply tiling and scale to shape texture
// Developer Intent: Apply a texture fill to a shape while safely handling missing image files and providing an automatic fallback texture.
// Use Cases: Load an external PNG as a texture for a rectangle shape and automatically switch to a built‑in texture when the file is absent. | Configure texture properties such as tiling and scaling without risking runtime crashes. | Save the workbook after applying the texture, capturing any errors that may occur during the save operation.
// AI Prompts: Generate C# code that sets a shape's texture fill from a file path in Aspose.Cells and uses a built‑in texture as a fallback if the file is missing. | Create robust error‑handling logic for applying a texture fill to a shape, covering FileNotFoundException and generic exceptions. | Refactor the example to extract texture loading into a reusable method with proper exception handling and fallback logic.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // Demonstrates how to add a rectangle shape to a workbook, load a PNG as a texture fill, verify the file exists, fall back to the built‑in WaterDroplets texture if it doesn't, and save the file with comprehensive error handling.
    public class ShapeTextureErrorHandlingDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Fatal error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape that will receive the texture fill
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 300);
            shape.Fill.FillType = FillType.Texture; // Enable texture fill

            // Path to the image that will be used as texture
            string imagePath = Path.Combine(Environment.CurrentDirectory, "texture.png");

            // Attempt to load the image data and assign it to the shape's texture fill
            try
            {
                if (!File.Exists(imagePath))
                {
                    // Throw a more descriptive exception if the file is missing
                    throw new FileNotFoundException($"Texture image file not found: {imagePath}");
                }

                // Read the image bytes and set them as the texture fill data
                byte[] imageData = File.ReadAllBytes(imagePath);
                shape.Fill.TextureFill.ImageData = imageData;

                // Optional: configure additional texture fill properties
                shape.Fill.TextureFill.IsTiling = true;
                shape.Fill.TextureFill.Scale = 0.5;
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.WriteLine($"Error: {fnfEx.Message}");
                // Fallback: use a built‑in texture type instead of a missing file
                shape.Fill.TextureFill.Type = TextureType.WaterDroplets;
                Console.WriteLine("Applied fallback built‑in texture type: WaterDroplets.");
            }
            catch (Exception ex)
            {
                // Catch any other unexpected errors
                Console.WriteLine($"Unexpected error: {ex.Message}");
                return;
            }

            // Save the workbook
            string outputPath = Path.Combine(Environment.CurrentDirectory, "ShapeTextureDemo.xlsx");
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
