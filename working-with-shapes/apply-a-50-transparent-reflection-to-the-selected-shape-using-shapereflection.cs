using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class ApplyReflectionDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Prepare a temporary PNG image (1x1 pixel) for the picture shape
            string tempImagePath = Path.Combine(Path.GetTempPath(), "tempShapeImage.png");
            // Base64 encoded PNG (transparent 1x1 pixel)
            const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(tempImagePath, pngBytes);

            // Ensure the temporary image file exists before adding the picture
            if (!File.Exists(tempImagePath))
                throw new FileNotFoundException("Temporary image file not found.", tempImagePath);

            // Add a picture shape using a file stream (compatible with recent Aspose.Cells versions)
            using (FileStream imgStream = File.OpenRead(tempImagePath))
            {
                // upperRow, upperColumn, lowerRow, lowerColumn define the shape's position
                Shape pictureShape = worksheet.Shapes.AddPicture(1, 0, 1, 0, imgStream);

                // Obtain the reflection effect object for the shape
                ReflectionEffect reflection = pictureShape.Reflection;
                if (reflection == null)
                    throw new InvalidOperationException("Reflection effect is not supported for this shape.");

                // Configure reflection properties
                reflection.Type = ReflectionEffectType.Custom;
                reflection.Transparency = 0.5;   // 50% transparency
                reflection.Size = 55;            // End alpha position (percentage)
                reflection.Blur = 0.5;           // Blur radius
                reflection.Distance = 0;        // Distance from the shape
            }

            // Save the workbook with the applied reflection effect
            string outputPath = "ShapeWithReflection.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}