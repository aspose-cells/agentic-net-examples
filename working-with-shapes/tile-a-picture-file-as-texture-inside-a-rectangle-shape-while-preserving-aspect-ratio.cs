// Title: Tile an Image as Texture in a Rectangle Shape with Aspect Ratio Preservation – Aspose.Cells for .NET (C#)
// Description: Shows how to insert a rectangle shape into an Excel worksheet using Aspose.Cells for .NET, apply a PNG texture fill, enable tiling, keep the picture’s original proportions via FillPictureType.StackAndScale, and save the workbook.
// Keywords: Aspose.Cells | C# | texture fill | image tiling | rectangle shape | FillPictureType.StackAndScale | preserve aspect ratio | Excel shape fill | Aspose.Cells example | IsTiling property
// Common Searches: Aspose.Cells tile image as texture in shape | C# preserve aspect ratio texture fill Aspose.Cells | How to use FillPictureType.StackAndScale with Aspose.Cells | Enable IsTiling for shape fill in Aspose.Cells .NET | Add rectangle shape with texture fill using Aspose.Cells
// Developer Intent: Apply a tiled picture texture to a rectangle shape in an Excel file while maintaining the image’s original proportions.
// Use Cases: Create a patterned background for a chart area by tiling a PNG texture within a rectangle. | Design a report header that repeats a logo texture without distortion. | Add a tiled watermark to a printable Excel flyer that keeps its aspect ratio.
// AI Prompts: Show how to change the tile scaling factor to 0.5 while keeping the image proportions in the provided Aspose.Cells code. | Provide a C# example that applies a tiled JPEG texture to a circular shape using Aspose.Cells for .NET. | Explain how to read the current TextureFill settings of a shape and switch the picture format type from StackAndScale to Stretch.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to insert a rectangle shape into an Excel worksheet using Aspose.Cells for .NET, apply a PNG texture fill, enable tiling, keep the picture’s original proportions via FillPictureType.StackAndScale, and save the workbook.
class TilePictureAsTexture
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape (row, column, offsetX, offsetY, width, height)
            Shape rectangle = worksheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 150);

            // Set the fill type of the shape to texture
            rectangle.Fill.FillType = FillType.Texture;

            // Path to the texture image
            string texturePath = "texture.png";

            if (File.Exists(texturePath))
            {
                // Load the picture file that will be used as texture
                byte[] imageData = File.ReadAllBytes(texturePath);

                // Configure the texture fill
                TextureFill textureFill = rectangle.Fill.TextureFill;
                textureFill.ImageData = imageData;          // assign image bytes
                textureFill.IsTiling = true;                // enable tiling
                textureFill.PictureFormatType = FillPictureType.StackAndScale; // preserve aspect ratio while tiling
                textureFill.Scale = 1.0;                    // overall scaling factor (1 = original size)
            }
            else
            {
                Console.WriteLine($"Texture file not found: {texturePath}. The rectangle will be saved without a texture.");
            }

            // Save the workbook
            string outputPath = "RectangleWithTiledTexture.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
