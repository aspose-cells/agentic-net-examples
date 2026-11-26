// Title: Tile a picture as a texture fill inside a rectangle shape while preserving aspect ratio with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# tile image as texture fill in a rectangle shape | preserve aspect ratio when using texture fill on shapes in Aspose.Cells | how to enable tiling for picture fill in Aspose.Cells worksheet | set TilePicOption scaling for texture fill in Aspose.Cells C# | apply JPEG as tiled texture to a shape using Aspose.Cells API
// Tags: texture fill tiling Aspose.Cells C# | rectangle shape texture fill Aspose.Cells | preserve aspect ratio texture fill Aspose.Cells | TilePicOption scaling Aspose.Cells | image data assignment texture fill Aspose.Cells | shape fill type texture Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace TextureTileDemo
{
    // The example reads a JPEG file into a byte array, creates a new workbook, adds a rectangle shape, sets its FillType to Texture, assigns the image data, enables tiling, configures TilePicOption with ScaleX and ScaleY of 1.0 to keep the original aspect ratio, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the image that will be used as a texture.
            string imagePath = "texture.jpg";

            if (!File.Exists(imagePath))
            {
                Console.WriteLine($"Error: Image file not found at path \"{imagePath}\".");
                return;
            }

            // Load image bytes.
            byte[] imageData;
            try
            {
                imageData = File.ReadAllBytes(imagePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading image file: {ex.Message}");
                return;
            }

            // Create a new workbook and get the first worksheet.
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet.
            // Parameters: upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, width, height (pixels)
            int shapeTopRow = 2;
            int shapeLeftColumn = 1;
            int shapeTopRowOffset = 0;
            int shapeLeftColumnOffset = 0;
            int shapeWidth = 300;   // pixels
            int shapeHeight = 200;  // pixels

            Shape rectangle = worksheet.Shapes.AddRectangle(
                shapeTopRow,
                shapeLeftColumn,
                shapeTopRowOffset,
                shapeLeftColumnOffset,
                shapeWidth,
                shapeHeight);

            // Configure the shape to use a texture fill.
            rectangle.Fill.FillType = FillType.Texture;
            TextureFill textureFill = rectangle.Fill.TextureFill;

            // Assign the image data to the texture fill.
            textureFill.ImageData = imageData;

            // Enable tiling.
            textureFill.IsTiling = true;

            // Set tile options with uniform scaling (no scaling to preserve original aspect ratio).
            TilePicOption tileOption = new TilePicOption
            {
                ScaleX = 1.0,
                ScaleY = 1.0,
                OffsetX = 0,
                OffsetY = 0
            };
            textureFill.TilePicOption = tileOption;

            // Optional: set transparency (0.0 = fully opaque).
            textureFill.Transparency = 0.0;

            // Save the workbook.
            string outputPath = "TiledTextureDemo.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
