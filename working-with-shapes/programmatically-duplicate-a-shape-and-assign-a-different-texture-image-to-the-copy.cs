// Title: Copy a Shape and Assign a Different Texture Image Using Aspose.Cells for .NET
// Description: This C# example creates a workbook, adds a rectangle shape with a PNG texture, duplicates the shape using AddCopy, replaces the copy's fill with another PNG or a built‑in texture, and saves the workbook.
// Keywords: Aspose.Cells | .NET | C# | duplicate shape | AddCopy | texture fill | custom image texture | Excel shape copy | shape fill type | fallback texture
// Common Searches: Aspose.Cells copy shape | AddCopy shape texture .NET | change shape fill after copy Aspose.Cells | set custom texture for Excel shape C# | fallback built‑in texture Aspose.Cells
// Developer Intent: Duplicate an existing worksheet shape and give the copy a different texture image.
// Use Cases: Design a reusable graphic element and reuse it with distinct backgrounds across a report. | Create a dashboard where each copied shape shows a different status icon. | Generate Excel files that automatically switch to a built‑in texture when a custom image is missing.
// AI Prompts: Generate C# code that copies any worksheet shape and applies a custom PNG texture using Aspose.Cells. | Show how to use AddCopy to position a duplicated shape and then change its FillType to a different image or built‑in texture. | Explain error handling for missing texture files when setting a shape's TextureFill in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, adds a rectangle shape with a PNG texture, duplicates the shape using AddCopy, replaces the copy's fill with another PNG or a built‑in texture, and saves the workbook.
    public class DuplicateShapeWithNewTexture
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add an initial rectangle shape to the worksheet
                // Parameters: upper left row, top offset, upper left column, left offset, width, height
                RectangleShape originalShape = worksheet.Shapes.AddRectangle(2, 0, 2, 0, 130, 130);

                // Optionally set a texture for the original shape (demonstration purpose)
                originalShape.Fill.FillType = FillType.Texture;
                TextureFill originalTexture = originalShape.Fill.TextureFill;

                // Load texture image if it exists
                string originalTexturePath = Path.Combine(Environment.CurrentDirectory, "texture1.png");
                if (File.Exists(originalTexturePath))
                {
                    originalTexture.ImageData = File.ReadAllBytes(originalTexturePath);
                    // No need to set Type; assigning ImageData makes it a custom texture
                }

                // Duplicate the rectangle shape to a new location using AddCopy
                // Parameters: source shape, top row, top offset, left column, left offset
                Shape copiedShape = worksheet.Shapes.AddCopy(originalShape, 7, 0, 7, 0);

                // Change the fill of the copied shape to use a different texture image
                copiedShape.Fill.FillType = FillType.Texture;
                TextureFill copiedTexture = copiedShape.Fill.TextureFill;

                // Load a different texture image file if it exists
                string newTexturePath = Path.Combine(Environment.CurrentDirectory, "texture2.png");
                if (File.Exists(newTexturePath))
                {
                    copiedTexture.ImageData = File.ReadAllBytes(newTexturePath);
                }
                else
                {
                    // Fallback to a built‑in texture type if the image file is missing
                    copiedTexture.Type = TextureType.WaterDroplets;
                }

                // Save the workbook
                string outputPath = Path.Combine(Environment.CurrentDirectory, "DuplicateShapeWithNewTexture.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            DuplicateShapeWithNewTexture.Run();
        }
    }
}
