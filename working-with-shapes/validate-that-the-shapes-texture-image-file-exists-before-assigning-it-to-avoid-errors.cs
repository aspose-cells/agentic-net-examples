// Title: Aspose.Cells .NET – Validate Texture Image Exists Before Assigning to a Shape
// Description: Shows how to create a workbook, confirm that a PNG texture file is present, add a rectangle shape, set its FillType to Texture, safely load the image bytes, optionally adjust tiling and scaling, and save the file—preventing FileNotFoundException.
// Keywords: Aspose.Cells | .NET | C# | shape texture fill | texture image validation | File.Exists | FillType.Texture | rectangle shape | error handling | image file check | Aspose.Cells Drawing
// Common Searches: Aspose.Cells check texture file before fill | C# Aspose.Cells shape texture error | Validate image path for shape fill Aspose.Cells | Prevent FileNotFoundException Aspose.Cells texture | How to apply texture fill safely in Aspose.Cells
// Developer Intent: Ensure the texture image file is present before assigning it to a shape’s texture fill to avoid runtime errors.
// Use Cases: Prevent FileNotFoundException when the texture image is missing. | Display a clear error message and abort the operation if the image cannot be found. | Apply a texture fill only when the image is available; otherwise fall back to a solid fill or skip the shape. | Reuse the validation logic across multiple worksheets or shape types.
// AI Prompts: Generate a C# helper method for Aspose.Cells that checks a texture image path and applies it to a given shape, returning a success flag. | Write code that logs a warning and continues workbook generation when the texture file is absent, while still adding the shape with a default fill. | Create a reusable .NET class that validates image existence and applies texture fills to multiple Aspose.Cells shapes with optional tiling and scaling parameters. | Provide a PowerShell script that scans a folder for missing texture images referenced in Aspose.Cells shape fills and reports them.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Shows how to create a workbook, confirm that a PNG texture file is present, add a rectangle shape, set its FillType to Texture, safely load the image bytes, optionally adjust tiling and scaling, and save the file—preventing FileNotFoundException.
class TextureShapeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the path to the texture image
        string imagePath = Path.Combine(Environment.CurrentDirectory, "texture.png");

        // Validate that the image file exists before using it
        if (!File.Exists(imagePath))
        {
            Console.WriteLine($"Error: Texture image file not found at '{imagePath}'.");
            return; // Exit to avoid assigning a non‑existent image
        }

        // Add a rectangle shape to the worksheet
        Shape rectangle = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 200);

        // Set the fill type to texture
        rectangle.Fill.FillType = FillType.Texture;

        // Assign the image data to the texture fill
        rectangle.Fill.TextureFill.ImageData = File.ReadAllBytes(imagePath);

        // Optional texture fill settings
        rectangle.Fill.TextureFill.IsTiling = true;
        rectangle.Fill.TextureFill.Scale = 0.5;

        // Save the workbook
        workbook.Save("TextureShapeDemo.xlsx");
        Console.WriteLine("Workbook saved successfully with texture applied.");
    }
}
