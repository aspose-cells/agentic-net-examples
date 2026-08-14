// Title: Validate texture image existence before applying a shape fill in Aspose.Cells for .NET (C#)
// Description: This C# example creates a workbook, adds a rectangle shape, checks that a PNG file (texture.png) is present in the current directory, and only then assigns the image bytes to the shape's texture fill with optional tiling. If the file is missing, a clear error is written and the process stops, preventing runtime exceptions.
// Keywords: Aspose.Cells | C# | .NET | shape texture fill | validate image file | File.Exists | rectangle shape | texture tiling | error handling | workbook save
// Common Searches: Aspose.Cells check image file before texture fill | C# prevent FileNotFoundException when using Shape.Fill.TextureFill | how to verify texture.png exists in Aspose.Cells example | apply texture fill to shape only if file is present Aspose.Cells | Aspose.Cells shape fill validation tutorial
// Developer Intent: Ensure a texture image is available before assigning it to a shape's fill to avoid exceptions.
// Use Cases: Guard against missing texture files when generating Excel reports with custom graphics. | Provide a user‑friendly error message and abort workbook creation if required assets are absent. | Enable tiling of a texture only after successful image loading.
// AI Prompts: Generate a reusable C# method that takes a Shape and an image path, verifies the file exists, and applies it as a texture fill using Aspose.Cells. | Write code that falls back to a solid color fill when a texture image cannot be found for a shape in Aspose.Cells. | Create a utility class for loading, validating, and optionally tiling texture images for shapes in Aspose.Cells for .NET.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// This C# example creates a workbook, adds a rectangle shape, checks that a PNG file (texture.png) is present in the current directory, and only then assigns the image bytes to the shape's texture fill with optional tiling. If the file is missing, a clear error is written and the process stops, preventing runtime exceptions.
class TextureValidationDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Define the path to the texture image file
        string imagePath = Path.Combine(Environment.CurrentDirectory, "texture.png");

        // Validate that the image file exists before using it
        if (!File.Exists(imagePath))
        {
            Console.WriteLine($"Error: Texture file not found at '{imagePath}'.");
            return; // Exit to avoid assigning a non‑existent image
        }

        // Add a rectangle shape to the worksheet
        Shape rectangle = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 150, 100);

        // Set the fill type of the shape to texture
        rectangle.Fill.FillType = FillType.Texture;

        // Assign the image data to the texture fill
        rectangle.Fill.TextureFill.ImageData = File.ReadAllBytes(imagePath);
        rectangle.Fill.TextureFill.IsTiling = true; // optional: enable tiling

        // Save the workbook
        workbook.Save("TextureValidated.xlsx");
        Console.WriteLine("Workbook saved successfully with validated texture.");
    }
}
