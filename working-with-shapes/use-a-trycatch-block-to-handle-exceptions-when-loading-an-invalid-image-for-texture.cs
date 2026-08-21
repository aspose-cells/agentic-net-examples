// Title: Handle Missing or Corrupt Texture Images with Try‑Catch in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to add a rectangle shape, set its FillType to Texture, and safely load image data using a try‑catch block. The code logs errors when the texture file is absent or unreadable and still saves the workbook.
// Keywords: Aspose.Cells texture fill exception handling | C# Aspose.Cells shape texture | try catch load image Aspose.Cells | invalid texture image handling | Aspose.Cells .NET error logging
// Common Searches: Aspose.Cells catch error loading texture image | C# shape texture fill missing file Aspose.Cells | how to handle invalid texture image in Aspose.Cells | exception handling for TextureFill in Aspose.Cells
// Developer Intent: The developer wants to load a texture image for a shape safely and ensure the workbook is saved even if the image cannot be loaded.
// Use Cases: Apply a texture to a shape while providing a graceful fallback when the image file is missing or corrupted. | Log detailed error information for troubleshooting texture loading issues. | Guarantee workbook persistence regardless of texture load success.
// AI Prompts: Create a C# example that sets a texture fill on an Aspose.Cells shape and includes try‑catch logic for missing or damaged image files. | Explain best practices for handling exceptions when assigning ImageData to a TextureFill object in Aspose.Cells and still saving the workbook. | Generate code that logs texture loading errors with context while using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureDemo
{
    // Demonstrates how to add a rectangle shape, set its FillType to Texture, and safely load image data using a try‑catch block. The code logs errors when the texture file is absent or unreadable and still saves the workbook.
    public class LoadInvalidTextureDemo
    {
        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to demonstrate texture fill
            Shape shape = worksheet.Shapes.AddRectangle(1, 0, 2, 5, 150, 300);
            shape.Fill.FillType = FillType.Texture;
            TextureFill textureFill = shape.Fill.TextureFill;

            try
            {
                string imagePath = "invalid_texture_image.png";

                // Ensure the image file exists before attempting to read it
                if (File.Exists(imagePath))
                {
                    byte[] imageData = File.ReadAllBytes(imagePath);
                    textureFill.ImageData = imageData;

                    // Set additional texture properties
                    textureFill.IsTiling = true;
                    textureFill.Scale = 0.8;
                }
                else
                {
                    Console.WriteLine($"Texture image file not found: {imagePath}");
                }
            }
            catch (Exception ex)
            {
                // Handle any errors that occur while loading the image
                Console.WriteLine($"Error loading texture image: {ex.Message}");
            }

            try
            {
                // Save the workbook (even if the texture load failed)
                workbook.Save("LoadInvalidTextureDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            LoadInvalidTextureDemo.Run();
        }
    }
}
