// Title: Use try‑catch to handle missing or corrupted image when applying a texture fill to a shape in Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells how to apply texture fill to shape with fallback when image file not found | handle exception loading image for Shape.Fill.TextureFill in Aspose.Cells .NET | use built‑in texture as backup if texture image load fails Aspose.Cells | read image bytes safely for texture fill Aspose.Cells C# try catch example | Aspose.Cells texture fill error handling for invalid image path
// Tags: shape texture fill exception handling | fallback built‑in texture Aspose.Cells | load image bytes safely .NET | Aspose.Cells texture fill from byte array | rectangle shape texture fill C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example creates a workbook, adds a rectangle shape, sets its fill type to texture, and attempts to load image data from a specified path. Image loading is wrapped in a try‑catch block; on failure the exception is logged and a built‑in blue tissue‑paper texture is applied as a fallback before saving the workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape that will demonstrate texture fill
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 100);
        shape.Fill.FillType = FillType.Texture;
        TextureFill textureFill = shape.Fill.TextureFill;

        // Path to an image that is intentionally invalid (non‑existent or corrupted)
        string imagePath = "invalid_image.jpg";

        try
        {
            // Attempt to read the image file into a byte array
            byte[] imageData = File.ReadAllBytes(imagePath);

            // Assign the image data to the texture fill
            textureFill.ImageData = imageData;
            Console.WriteLine("Image loaded successfully and applied as texture.");
        }
        catch (Exception ex)
        {
            // Handle any exception that occurs while loading the image
            Console.WriteLine($"Error loading image for texture: {ex.Message}");

            // Fallback: use a built‑in texture type so the workbook can still be saved
            textureFill.Type = TextureType.BlueTissuePaper;
            textureFill.IsTiling = true;
            Console.WriteLine("Fallback to built‑in texture applied.");
        }

        // Save the workbook to verify the result
        string outputFile = "TextureFillDemo.xlsx";
        workbook.Save(outputFile);
        Console.WriteLine($"Workbook saved to \"{outputFile}\".");
    }
}
