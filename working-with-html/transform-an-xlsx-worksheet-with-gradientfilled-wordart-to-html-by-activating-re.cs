using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class GradientWordArtToHtml
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape with a preset gradient style (e.g., WordArtStyle6)
        // Parameters: style, text, topRow, top, leftColumn, left, height, width
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle6, // Gradient Fill - Gray
            "Gradient WordArt",
            2,   // top row index
            10,  // vertical offset in pixels
            2,   // left column index
            10,  // horizontal offset in pixels
            100, // height in pixels
            400  // width in pixels
        );

        // Customize the gradient fill
        wordArt.Fill.FillType = FillType.Gradient;
        GradientFill gradient = wordArt.Fill.GradientFill;
        gradient.SetTwoColorGradient(
            Color.LightGray,          // first color
            Color.DarkGray,           // second color
            GradientStyleType.Horizontal,
            1                         // variant
        );

        // Prepare HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Save the workbook as HTML
        workbook.Save("GradientWordArt.html", htmlOptions);
    }
}