using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add WordArt with a preset style that already contains a gradient fill and reflection
        // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,
            "Reflection Demo",
            2, 0, 2, 0, 300, 100);

        // Verify the shape is WordArt before applying reflection settings
        if (wordArt.IsWordArt)
        {
            // Access the reflection effect of the WordArt shape
            ReflectionEffect reflection = wordArt.Reflection;

            // Customize reflection properties as needed
            reflection.Type = ReflectionEffectType.FullReflectionTouching;
            reflection.Transparency = 0.3;   // 30% transparent at start
            reflection.Size = 80;           // 80% size of the original shape
            reflection.Blur = 5;            // Blur radius in points
            reflection.Distance = 5;        // Distance from the shape in points
        }

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Export images (including gradient fills) as Base64 to keep them visible in HTML
            ExportImagesAsBase64 = true,
            // Ensure all visible shapes are exported
            IgnoreInvisibleShapes = false
        };

        // Save the workbook as an HTML file with the specified options
        workbook.Save("WordArtReflection.html", htmlOptions);
    }
}