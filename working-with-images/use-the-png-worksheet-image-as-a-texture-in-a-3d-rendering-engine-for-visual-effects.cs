using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsTextureDemo
{
    class Program
    {
        static void Main()
        {
            // 1. Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Row 1");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["A3"].PutValue("Row 2");
            sheet.Cells["B3"].PutValue(456);

            // 2. Render the first worksheet page to a PNG image in memory
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                ImageType = ImageType.Png,
                OnePagePerSheet = true
            };
            SheetRender sheetRender = new SheetRender(sheet, renderOptions);
            byte[] pngBytes;
            using (MemoryStream pngStream = new MemoryStream())
            {
                // Use the provided ToImage(int, Stream) method
                sheetRender.ToImage(0, pngStream);
                pngBytes = pngStream.ToArray(); // Capture the rendered PNG bytes
            }

            // 3. Add a rectangle shape that will use the rendered PNG as its texture
            // Parameters: upper left row, upper left column, upper left offsetX, offsetY, width, height
            Shape textureShape = sheet.Shapes.AddRectangle(5, 0, 0, 0, 200, 150);
            // Set the fill type to texture
            textureShape.Fill.FillType = FillType.Texture;

            // 4. Apply the rendered PNG as the texture image data
            TextureFill textureFill = textureShape.Fill.TextureFill;
            textureFill.ImageData = pngBytes;          // Use the PNG bytes as texture
            textureFill.IsTiling = true;               // Tile the texture for demonstration
            textureFill.Transparency = 0.1;            // Slight transparency
            textureFill.Scale = 0.8;                   // Scale the texture

            // 5. Optionally configure 3D format to enhance visual effect
            textureShape.ThreeDFormat.ExtrusionHeight = 20;
            textureShape.ThreeDFormat.Material = PresetMaterialType.Metal;

            // 6. Save the workbook with the textured shape
            workbook.Save("WorkbookWithTextureShape.xlsx");

            Console.WriteLine("Workbook created. The worksheet was rendered to PNG and applied as a texture to a shape.");
        }
    }
}