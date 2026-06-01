using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class TextureTilingComparison
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Load texture image bytes (ensure the file exists)
            string imagePath = "texture.png";
            if (!File.Exists(imagePath))
                throw new FileNotFoundException($"Texture file not found: {imagePath}");

            byte[] textureData = File.ReadAllBytes(imagePath);

            // -------------------------------------------------
            // Shape with tiled texture (IsTiling = true)
            // -------------------------------------------------
            Shape tiledShape = sheet.Shapes.AddRectangle(1, 0, 1, 0, 200, 200);
            tiledShape.Fill.FillType = FillType.Texture;
            tiledShape.Fill.TextureFill.ImageData = textureData;
            tiledShape.Fill.TextureFill.IsTiling = true;               // Tile the picture
            tiledShape.Fill.TextureFill.Type = TextureType.Unknown;    // Use the supplied image data

            // -------------------------------------------------
            // Shape with stretched texture (IsTiling = false)
            // -------------------------------------------------
            Shape stretchedShape = sheet.Shapes.AddRectangle(1, 0, 1, 250, 200, 200);
            stretchedShape.Fill.FillType = FillType.Texture;
            stretchedShape.Fill.TextureFill.ImageData = textureData;
            stretchedShape.Fill.TextureFill.IsTiling = false;          // Stretch the picture
            stretchedShape.Fill.TextureFill.Type = TextureType.Unknown;

            // Save the workbook to compare the two visual results
            workbook.Save("TextureTilingComparison.xlsx");
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"File error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}