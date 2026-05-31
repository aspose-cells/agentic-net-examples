using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class TextureTileDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add a rectangle shape (row, column, row offset, column offset, width, height)
            Shape rect = sheet.Shapes.AddRectangle(2, 2, 0, 0, 300, 200);

            // Set the fill type of the shape to texture
            rect.Fill.FillType = FillType.Texture;
            TextureFill texFill = rect.Fill.TextureFill;

            // Load the picture file to be used as texture if it exists
            string texturePath = "texture.png"; // adjust path as needed
            if (File.Exists(texturePath))
            {
                byte[] imgData = File.ReadAllBytes(texturePath);
                texFill.ImageData = imgData;

                // Enable tiling of the picture
                texFill.IsTiling = true;

                // Preserve aspect ratio while tiling
                texFill.PictureFormatType = FillPictureType.StackAndScale;

                // Optional: configure tile options (scale and alignment)
                texFill.TilePicOption = new TilePicOption
                {
                    ScaleX = 100,               // 100% horizontal scale
                    ScaleY = 100,               // 100% vertical scale
                    AlignmentType = RectangleAlignmentType.Center
                };
            }
            else
            {
                Console.WriteLine($"Texture file '{texturePath}' not found. Skipping texture fill.");
            }

            // Save the workbook with the textured rectangle
            string outputPath = "TextureTileDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}