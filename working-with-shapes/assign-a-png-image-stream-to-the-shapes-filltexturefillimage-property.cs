using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a rectangle shape to the worksheet
        Shape shape = worksheet.Shapes.AddRectangle(0, 0, 0, 0, 200, 100);

        // Set the fill type of the shape to texture
        shape.Fill.FillType = FillType.Texture;

        // Obtain the TextureFill object from the shape's fill format
        TextureFill textureFill = shape.Fill.TextureFill;

        // Load a PNG image from a file stream into a byte array
        byte[] pngData;
        using (FileStream fileStream = new FileStream("sample.png", FileMode.Open, FileAccess.Read))
        using (MemoryStream memoryStream = new MemoryStream())
        {
            fileStream.CopyTo(memoryStream);
            pngData = memoryStream.ToArray();
        }

        // Assign the PNG image data to the texture fill
        textureFill.ImageData = pngData;

        // Save the workbook with the textured shape
        workbook.Save("ShapeWithTexture.xlsx");
    }
}