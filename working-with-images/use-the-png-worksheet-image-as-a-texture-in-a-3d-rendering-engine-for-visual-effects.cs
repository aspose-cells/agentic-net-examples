using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // 1. Create a new workbook and add some sample data.
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apples");
        worksheet.Cells["B2"].PutValue(150);
        worksheet.Cells["A3"].PutValue("Oranges");
        worksheet.Cells["B3"].PutValue(200);

        // 2. Render the worksheet to a PNG image in memory using SheetRender.
        ImageOrPrintOptions renderOptions = new ImageOrPrintOptions();
        renderOptions.ImageType = ImageType.Png;               // PNG output
        SheetRender sheetRender = new SheetRender(worksheet, renderOptions);

        byte[] worksheetPng;
        using (MemoryStream pngStream = new MemoryStream())
        {
            // Render first page (index 0) to the stream – follows the provided rule.
            sheetRender.ToImage(0, pngStream);
            worksheetPng = pngStream.ToArray();                // Capture image bytes
        }

        // 3. Add a shape that will use the rendered worksheet image as a texture.
        //    The shape is a rectangle positioned at row 5, column 2.
        Shape texturedShape = worksheet.Shapes.AddRectangle(5, 2, 5, 2, 250, 150);
        texturedShape.Fill.FillType = FillType.Texture;        // Enable texture fill

        // 4. Configure the texture fill with the image data from step 2.
        TextureFill textureFill = texturedShape.Fill.TextureFill;
        textureFill.ImageData = worksheetPng;                  // Set the PNG as texture
        textureFill.IsTiling = true;                          // Tile the texture
        textureFill.Scale = 0.9;                               // Slightly shrink the texture

        // 5. Apply 3‑D formatting to enhance visual effects.
        texturedShape.ThreeDFormat.Material = PresetMaterialType.Metal; // Metallic look
        texturedShape.ThreeDFormat.ExtrusionHeight = 20;                // Give depth
        texturedShape.ThreeDFormat.RotationX = 25;                     // Tilt X
        texturedShape.ThreeDFormat.RotationY = 15;                     // Tilt Y

        // 6. Save the workbook with the textured 3‑D shape.
        workbook.Save("WorkbookWith3DTextureShape.xlsx");
    }
}