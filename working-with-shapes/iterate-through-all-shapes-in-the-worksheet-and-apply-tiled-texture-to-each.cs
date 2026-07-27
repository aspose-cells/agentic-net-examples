// Title: C# – Apply a Tiled Canvas Texture Fill to All Shapes in an Aspose.Cells Worksheet
// Description: Creates a workbook, adds rectangle, oval and textbox shapes, then iterates through every shape to set the fill type to Texture, selects the built‑in Canvas texture, enables tiling, and customizes TilePicOption (scale and offset) before saving as TextureFilledShapes.xlsx.
// Keywords: Aspose.Cells C# texture fill | shape tiling Aspose.Cells | TextureFill Canvas Aspose.Cells | TilePicOption example | apply texture to all shapes | Excel shape fill C# | Aspose.Cells shape formatting
// Common Searches: C# Aspose.Cells apply tiled texture to shapes | how to set Canvas texture fill for worksheet shapes | iterate shapes and enable tiling Aspose.Cells | TilePicOption usage in Aspose.Cells C# | texture fill all shapes Excel Aspose
// Developer Intent: Programmatically give every shape in a worksheet a repeating Canvas texture with custom scaling and offset.
// Use Cases: Generate reports where all diagram elements share a consistent tiled background. | Create visual templates with automatically textured shapes, avoiding manual formatting. | Apply custom scaling and offset to texture fills for precise visual patterns in Excel files.
// AI Prompts: Write C# code using Aspose.Cells that loops through all worksheet shapes and applies a tiled Canvas texture with specific ScaleX, ScaleY, OffsetX, and OffsetY values. | Show how to use TextureFill and TilePicOption classes to set a repeating texture for shapes added to an Excel workbook. | Provide an example that adds multiple shapes to a worksheet, enables tiling for their texture fills, and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureDemo
{
    // Creates a workbook, adds rectangle, oval and textbox shapes, then iterates through every shape to set the fill type to Texture, selects the built‑in Canvas texture, enables tiling, and customizes TilePicOption (scale and offset) before saving as TextureFilledShapes.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Example: add a few shapes to demonstrate the texture application
            worksheet.Shapes.AddRectangle(2, 0, 2, 0, 100, 200);
            worksheet.Shapes.AddOval(5, 0, 5, 0, 80, 120);
            worksheet.Shapes.AddTextBox(8, 0, 8, 0, 150, 60);

            // Iterate through all shapes in the worksheet
            for (int i = 0; i < worksheet.Shapes.Count; i++)
            {
                Shape shape = worksheet.Shapes[i];

                // Set the fill type to texture
                shape.Fill.FillType = FillType.Texture;

                // Get the TextureFill object
                TextureFill textureFill = shape.Fill.TextureFill;

                // Choose a built‑in texture type (any type can be used)
                textureFill.Type = TextureType.Canvas;

                // Enable tiling
                textureFill.IsTiling = true;

                // Optional: configure tile options (scale, offset, etc.)
                TilePicOption tileOptions = new TilePicOption
                {
                    ScaleX = 0.5,   // 50% width scaling
                    ScaleY = 0.5,   // 50% height scaling
                    OffsetX = 10,   // horizontal offset
                    OffsetY = 10    // vertical offset
                };
                textureFill.TilePicOption = tileOptions;
            }

            // Save the workbook with the applied texture fills
            workbook.Save("TextureFilledShapes.xlsx");
        }
    }
}
