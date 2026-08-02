// Title: C# – Aspose.Cells Texture Fill with Try‑Catch Exception Handling for Invalid Images
// Description: Demonstrates how to add a rectangle shape to a workbook, set its fill type to Texture, and safely load an image file using try‑catch blocks. The example handles missing or corrupt image files, reports errors during image loading, and catches exceptions when saving the workbook or at the application entry point.
// Keywords: Aspose.Cells | C# texture fill | shape texture exception handling | try catch image load | invalid texture image | Aspose.Cells error handling
// Common Searches: Aspose.Cells texture fill exception handling C# | how to catch errors loading shape texture image Aspose.Cells | C# try catch invalid image Aspose.Cells shape | Aspose.Cells missing texture file handling
// Developer Intent: The developer wants to apply a texture fill to a shape while gracefully handling file‑related and I/O errors.
// Use Cases: Load a texture image from disk and apply it to a shape, with clear messages for missing or corrupted files. | Prevent application crashes by catching exceptions during workbook saving (e.g., permission or path issues). | Provide a top‑level safety net that logs unexpected failures and ensures a clean exit.
// AI Prompts: Generate C# code using Aspose.Cells that applies a texture fill to a shape and includes robust try‑catch handling for missing or invalid image files. | Show how to log detailed exception information (type, stack trace) when loading a texture image for a shape in Aspose.Cells. | Recommend best‑practice error‑handling patterns for Aspose.Cells texture fills, including specific exception types to catch and logging strategies.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a rectangle shape to a workbook, set its fill type to Texture, and safely load an image file using try‑catch blocks. The example handles missing or corrupt image files, reports errors during image loading, and catches exceptions when saving the workbook or at the application entry point.
public class TextureFillExceptionDemo
{
    public static void Run()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to demonstrate texture fill
        Shape shape = worksheet.Shapes.AddRectangle(0, 0, 0, 0, 200, 100);
        shape.Fill.FillType = FillType.Texture; // Set fill type to texture
        TextureFill textureFill = shape.Fill.TextureFill;

        // Load texture image if the file exists
        string texturePath = "invalid_texture.png";
        if (File.Exists(texturePath))
        {
            try
            {
                byte[] imageData = File.ReadAllBytes(texturePath);
                textureFill.ImageData = imageData;
                Console.WriteLine("Texture image loaded successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading texture image: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine($"Texture file not found: {texturePath}");
        }

        // Save the workbook
        try
        {
            workbook.Save("TextureFillExceptionDemo.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }

    // Entry point for the application
    public static void Main()
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unhandled exception: {ex.Message}");
        }
    }
}
