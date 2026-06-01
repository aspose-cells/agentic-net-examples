using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule: create)
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample shapes to demonstrate the operation
        worksheet.Shapes.AddRectangle(1, 0, 1, 0, 100, 200);
        worksheet.Shapes.AddOval(5, 0, 5, 0, 80, 120);
        worksheet.Shapes.AddLine(2, 0, 2, 100, 200, 0);

        // Iterate through all shapes in the worksheet
        foreach (Shape shape in worksheet.Shapes)
        {
            // Set the fill type to texture
            shape.Fill.FillType = FillType.Texture;

            // Get the TextureFill object
            TextureFill textureFill = shape.Fill.TextureFill;

            // Choose a built‑in texture type (any type can be used)
            textureFill.Type = TextureType.BlueTissuePaper;

            // Enable tiling of the texture
            textureFill.IsTiling = true;

            // Optional: configure tile picture options (scale, offset, etc.)
            textureFill.TilePicOption = new TilePicOption
            {
                ScaleX = 0.5,   // 50% width scaling
                ScaleY = 0.5,   // 50% height scaling
                OffsetX = 0,
                OffsetY = 0
            };
        }

        // Save the workbook with the applied texture fills (lifecycle rule: save)
        workbook.Save("AllShapesTiledTexture.xlsx");
    }
}