using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // --------------------------------------------------------------------
                // Add an original rectangle shape and apply a texture fill to it
                // --------------------------------------------------------------------
                Shape originalShape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 150);
                originalShape.Fill.FillType = FillType.Texture;                     // enable texture fill
                TextureFill originalTexture = originalShape.Fill.TextureFill;       // get the TextureFill object

                // Load the first texture image from file and assign it (if it exists)
                string texturePath1 = "texture1.png";
                if (File.Exists(texturePath1))
                {
                    originalTexture.ImageData = File.ReadAllBytes(texturePath1);
                }
                else
                {
                    Console.WriteLine($"Warning: Texture file '{texturePath1}' not found. Skipping texture assignment.");
                }
                originalTexture.Type = TextureType.WaterDroplets;                  // optional built‑in texture type

                // --------------------------------------------------------------------
                // Duplicate the shape using ShapeCollection.AddCopy
                // --------------------------------------------------------------------
                // Parameters: source shape, top row, top offset (pixels), left column, left offset (pixels)
                Shape copiedShape = worksheet.Shapes.AddCopy(originalShape, 5, 0, 5, 0);

                // --------------------------------------------------------------------
                // Change the texture of the copied shape to a different image
                // --------------------------------------------------------------------
                copiedShape.Fill.FillType = FillType.Texture;                      // ensure texture fill is enabled
                TextureFill copiedTexture = copiedShape.Fill.TextureFill;          // get the TextureFill for the copy

                // Load a different texture image and assign it (if it exists)
                string texturePath2 = "texture2.png";
                if (File.Exists(texturePath2))
                {
                    copiedTexture.ImageData = File.ReadAllBytes(texturePath2);
                }
                else
                {
                    Console.WriteLine($"Warning: Texture file '{texturePath2}' not found. Skipping texture assignment.");
                }
                copiedTexture.Type = TextureType.Granite;                          // optional built‑in texture type

                // --------------------------------------------------------------------
                // Save the workbook with the original and duplicated shapes
                // --------------------------------------------------------------------
                string outputFile = "DuplicatedShapeWithDifferentTexture.xlsx";
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}