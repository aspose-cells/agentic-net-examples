using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsTextureTilingDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Add a rectangle shape to the worksheet
            // Parameters: upper left row, upper left column, upper left row offset, upper left column offset, width, height
            Shape shape = worksheet.Shapes.AddRectangle(2, 1, 0, 0, 200, 100);

            // Set the fill type of the shape to texture
            shape.Fill.FillType = FillType.Texture;

            // Access the TextureFill object
            TextureFill textureFill = shape.Fill.TextureFill;

            // Optionally set a built‑in texture type (e.g., BlueTissuePaper)
            textureFill.Type = TextureType.BlueTissuePaper;

            // Enable tiling for the texture fill
            textureFill.IsTiling = true;

            // Save the workbook to a file
            workbook.Save("TextureTilingDemo.xlsx");

            Console.WriteLine("Workbook saved with texture tiling enabled.");
        }
    }
}