// Title: How to check for a texture image file before assigning it to a shape’s TextureFill in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# verify texture image file exists before setting shape TextureFill | C# Aspose.Cells shape texture fill with file existence validation | prevent missing texture file error when applying texture fill to shape in Aspose.Cells | load texture image into shape TextureFill only if file is present using Aspose.Cells | Aspose.Cells .NET texture fill shape error handling for absent image file
// Tags: shape texture fill file validation | Aspose.Cells TextureFill ImageData assignment | C# check image file existence before applying texture | Aspose.Cells shape fill error handling | texture image handling in Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeTextureValidationDemo
{
    // The example creates a workbook, adds a rectangle shape, sets its fill type to Texture, and safely assigns an image to the shape’s TextureFill only after confirming the image file exists, with optional tiling, scaling, and transparency settings, while handling missing‑file scenarios gracefully.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the texture image file.
            // If a command‑line argument is supplied, use it; otherwise use a default file name.
            string texturePath = args.Length > 0 ? args[0] : "texture.png";

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to demonstrate texture fill.
            // Parameters: upper left row, upper left column, upper left offset in pixels,
            // upper left offset in pixels, width in pixels, height in pixels.
            Shape shape = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 200, 150);

            // Set the shape's fill type to Texture.
            shape.Fill.FillType = FillType.Texture;

            // Obtain the TextureFill object.
            TextureFill textureFill = shape.Fill.TextureFill;

            // Validate that the texture file exists before assigning it.
            if (File.Exists(texturePath))
            {
                try
                {
                    // Read the image bytes and assign to ImageData.
                    byte[] imageBytes = File.ReadAllBytes(texturePath);
                    textureFill.ImageData = imageBytes;

                    // Optional: configure additional texture properties.
                    textureFill.IsTiling = true;   // Tile the picture as texture.
                    textureFill.Scale = 0.8;       // Scale the texture.
                    textureFill.Transparency = 0.1; // Slight transparency.

                    Console.WriteLine($"Texture image '{texturePath}' loaded and applied to the shape.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error reading texture file '{texturePath}': {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine($"Warning: Texture file '{texturePath}' does not exist. The shape will retain default fill.");
            }

            // Save the workbook.
            string outputFile = "ShapeWithTextureDemo.xlsx";
            try
            {
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully as '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
