// Title: How to add error handling for missing image files when applying a texture fill to a shape using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set shape texture from external image with fallback if file is missing | Load image bytes into TextureFill.ImageData and catch exceptions in Aspose.Cells | Apply built‑in texture type when texture image cannot be read in Aspose.Cells workbook | Check File.Exists before assigning TextureFill.ImageData in a C# Aspose.Cells project | Configure tiling, scaling and transparency for a shape texture while handling missing files in Aspose.Cells
// Tags: Aspose.Cells texture fill error handling | TextureFill.ImageData file existence verification | default texture type Aspose.Cells | shape rectangle texture fill C# | tiling scaling transparency texture Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, adds a rectangle shape, sets its fill to texture, attempts to load image bytes from a specified path, falls back to a built‑in texture when the file is missing or unreadable, configures tiling, scaling and transparency, and saves the file.
class Program
{
    static void Main(string[] args)
    {
        // Determine the image file path. Use the first command‑line argument if provided,
        // otherwise fall back to a default file name in the current directory.
        string imagePath = args.Length > 0 ? args[0] : "texture.png";

        // Create a new workbook and get the first worksheet.
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape that will receive the texture fill.
        Shape shape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 300);

        // Set the fill type to texture and obtain the TextureFill object.
        shape.Fill.FillType = FillType.Texture;
        TextureFill textureFill = shape.Fill.TextureFill;

        // Attempt to load the image file and assign it to the texture fill.
        // If the file does not exist or cannot be read, fall back to a built‑in texture.
        if (File.Exists(imagePath))
        {
            try
            {
                textureFill.ImageData = File.ReadAllBytes(imagePath);
                Console.WriteLine($"Image data successfully loaded from \"{imagePath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading image file \"{imagePath}\": {ex.Message}");
                ApplyBuiltInTexture(textureFill);
            }
        }
        else
        {
            Console.WriteLine($"Image file \"{imagePath}\" not found.");
            ApplyBuiltInTexture(textureFill);
        }

        // Configure additional texture properties.
        textureFill.IsTiling = true;          // Tile the picture across the shape.
        textureFill.Scale = 0.8;              // Scale the picture to 80 % of its original size.
        textureFill.Transparency = 0.1;       // 10 % transparent.

        // Save the workbook.
        string outputFile = "ShapeTextureDemo.xlsx";
        try
        {
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to \"{outputFile}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to save workbook: {ex.Message}");
        }
    }

    /// <param name="textureFill">The TextureFill instance to configure.</param>
    private static void ApplyBuiltInTexture(TextureFill textureFill)
    {
        // Choose any built‑in texture type; here we use BlueTissuePaper.
        textureFill.Type = TextureType.BlueTissuePaper;
        Console.WriteLine("Applied built‑in texture type \"BlueTissuePaper\" as fallback.");
    }
}
