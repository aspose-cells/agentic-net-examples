using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class GradientWordArtToHtml
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape
        // Parameters: preset style, text, topRow, top, leftColumn, left, height, width
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle1,   // base style (will be overridden)
            "Gradient WordArt",
            2, 0,                               // top row and vertical offset
            2, 0,                               // left column and horizontal offset
            100, 400);                          // height and width

        // Set the fill type to gradient
        wordArt.Fill.FillType = FillType.Gradient;

        // Apply a two‑color gradient (red to yellow, horizontal)
        wordArt.Fill.SetTwoColorGradient(
            Color.Red,                         // first color
            Color.Yellow,                      // second color
            GradientStyleType.Horizontal,      // gradient style
            1);                                // variant

        // Configure HTML save options (optional settings)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExportImagesAsBase64 = true;   // embed images (including WordArt) as Base64

        // Save the workbook as HTML, preserving the gradient WordArt rendering
        workbook.Save("GradientWordArt.html", htmlOptions);
    }
}