// Title: C# – Handle missing image file when applying texture fill to a shape in Aspose.Cells
// Description: Demonstrates how to add a rectangle shape to a workbook, set its FillType to Texture, load image data from a file, and use try‑catch to apply a built‑in fallback texture (BlueTissuePaper) if the file is absent or unreadable, then save the workbook.
// Keywords: Aspose.Cells texture fill C# | shape texture error handling | fallback texture Aspose.Cells | FileNotFoundException shape fill | Aspose.Cells missing image | C# workbook shape texture | Aspose.Cells built‑in textures
// Common Searches: Aspose.Cells catch missing image for shape texture | C# set fallback texture when image file not found | How to use built‑in texture in Aspose.Cells | Error handling for TextureFill.ImageData | Apply default texture if external PNG missing Aspose.Cells
// Developer Intent: Add resilient code that substitutes a built‑in texture when the external image for a shape’s texture cannot be loaded.
// Use Cases: Load a PNG as a texture for a rectangle and automatically switch to TextureType.BlueTissuePaper if the file does not exist. | Log a clear console warning while allowing the workbook generation to continue. | Guarantee successful workbook saving regardless of texture image availability.
// AI Prompts: Generate C# Aspose.Cells code that applies a texture fill to a shape and falls back to a built‑in texture when the image file is missing. | Create an example that validates an image path before assigning it to TextureFill.ImageData and uses TextureType.BlueTissuePaper as a default. | Write a reusable method in C# that sets a shape’s texture from a file path with try‑catch handling for FileNotFoundException.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangle shape to a workbook, set its FillType to Texture, load image data from a file, and use try‑catch to apply a built‑in fallback texture (BlueTissuePaper) if the file is absent or unreadable, then save the workbook.
class ShapeTextureExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a rectangle shape
        Shape shape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 300);
        shape.Fill.FillType = FillType.Texture;
        TextureFill textureFill = shape.Fill.TextureFill;

        string imagePath = "texture.png";

        try
        {
            // Try to load image data from file
            byte[] imageData = File.ReadAllBytes(imagePath);
            textureFill.ImageData = imageData;
        }
        catch (FileNotFoundException)
        {
            Console.WriteLine($"Image file '{imagePath}' not found. Applying fallback texture.");
            // Use a built‑in texture as a fallback
            textureFill.Type = TextureType.BlueTissuePaper;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading texture: {ex.Message}");
        }

        // Save the workbook
        workbook.Save("ShapeTextureDemo.xlsx");
    }
}
