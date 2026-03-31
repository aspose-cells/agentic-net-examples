using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsWordArtHtmlDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a WordArt shape with a gradient preset style (e.g., WordArtStyle6)
            // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
            Shape wordArt = sheet.Shapes.AddWordArt(
                PresetWordArtStyle.WordArtStyle6, // Gradient Fill - Gray
                "Gradient WordArt",
                2, 0,   // row, top offset
                2, 0,   // column, left offset
                100,    // height
                400);   // width

            // Configure HTML save options to preserve shape rendering (including gradients)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export shapes as images encoded in Base64 to keep visual fidelity
                ExportImagesAsBase64 = true,

                // Ensure that invisible shapes are not ignored (optional, default is false)
                IgnoreInvisibleShapes = false
            };

            // Save the workbook as HTML (lifecycle: save)
            workbook.Save("WordArtGradient.html", htmlOptions);
        }
    }
}