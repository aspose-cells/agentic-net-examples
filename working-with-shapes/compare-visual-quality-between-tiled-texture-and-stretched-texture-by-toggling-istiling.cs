// Title: Aspose.Cells .NET: Compare Tiled vs Stretched Shape Texture Fill Using IsTiling
// Description: This example creates a new workbook, adds two rectangle shapes, loads a PNG file as texture data, and applies it with FillType.Texture. The first shape uses TextureFill.IsTiling = true for a repeating pattern, the second uses IsTiling = false for a stretched image. The workbook is saved so you can see the visual difference. The code also handles a missing texture file gracefully.
// Keywords: Aspose.Cells | .NET | C# | shape texture fill | IsTiling | tiled texture | stretched texture | FillType.Texture | Excel shape fill | Aspose.Cells example | rectangle shape texture
// Common Searches: How to apply tiled texture fill to a shape in Aspose.Cells | IsTiling true vs false effect on shape fill | Aspose.Cells compare tiled and stretched texture | C# code for texture fill with IsTiling property | Aspose.Cells shape fill scaling and tiling
// Developer Intent: Show how toggling the IsTiling property changes the rendering of a texture fill on Excel shapes.
// Use Cases: Create a repeating background pattern in a worksheet using a tiled texture. | Display a logo or banner without repetition by stretching the texture across a shape. | Generate a side‑by‑side visual test to decide which fill mode best fits a design. | Programmatically switch between tiled and stretched fills based on user preferences.
// AI Prompts: Write C# code that adds two rectangle shapes with the same PNG texture, one with IsTiling true and one with IsTiling false, then saves the workbook. | Explain the impact of the IsTiling property on texture rendering in Aspose.Cells and suggest when to use tiled versus stretched fills. | Modify the sample to set TextureFill.Scale to 0.5 for the tiled shape and 2.0 for the stretched shape, and describe the resulting appearance.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsExamples
{
    // This example creates a new workbook, adds two rectangle shapes, loads a PNG file as texture data, and applies it with FillType.Texture. The first shape uses TextureFill.IsTiling = true for a repeating pattern, the second uses IsTiling = false for a stretched image. The workbook is saved so you can see the visual difference. The code also handles a missing texture file gracefully.
    public class CompareTiledAndStretchedTexture
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Load texture image data (ensure the file exists)
                string texturePath = "texture.png";
                byte[] textureData = null;

                if (File.Exists(texturePath))
                {
                    textureData = File.ReadAllBytes(texturePath);
                }
                else
                {
                    Console.WriteLine($"Warning: Texture file '{texturePath}' not found. Shapes will be created without texture.");
                }

                // -------------------------------------------------
                // Shape 1: Tiled texture (IsTiling = true)
                // -------------------------------------------------
                Shape tiledShape = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 150);
                if (textureData != null)
                {
                    tiledShape.Fill.FillType = FillType.Texture;               // Use texture fill
                    tiledShape.Fill.TextureFill.ImageData = textureData;       // Set image data
                    tiledShape.Fill.TextureFill.IsTiling = true;               // Enable tiling
                    tiledShape.Fill.TextureFill.Scale = 1.0;                    // No additional scaling
                }

                // -------------------------------------------------
                // Shape 2: Stretched texture (IsTiling = false)
                // -------------------------------------------------
                Shape stretchedShape = worksheet.Shapes.AddRectangle(1, 0, 1, 200, 200, 150);
                if (textureData != null)
                {
                    stretchedShape.Fill.FillType = FillType.Texture;           // Use texture fill
                    stretchedShape.Fill.TextureFill.ImageData = textureData;   // Same image data
                    stretchedShape.Fill.TextureFill.IsTiling = false;          // Disable tiling (stretched)
                    stretchedShape.Fill.TextureFill.Scale = 1.0;                // No additional scaling
                }

                // Output status to console
                Console.WriteLine("Tiled shape IsTiling: " + (textureData != null ? tiledShape.Fill.TextureFill.IsTiling.ToString() : "N/A"));
                Console.WriteLine("Stretched shape IsTiling: " + (textureData != null ? stretchedShape.Fill.TextureFill.IsTiling.ToString() : "N/A"));

                // Save the workbook to visualize the result
                string outputPath = "CompareTiledAndStretchedTexture.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CompareTiledAndStretchedTexture.Run();
        }
    }
}
