using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class EnableWordArtGradientHtml
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape with a gradient preset style
        // Parameters: style, text, upperLeftRow, upperLeftColumn, upperLeftRowOffset, upperLeftColumnOffset, height, width
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle6,
            "Gradient WordArt",
            2, // upper left row
            2, // upper left column
            0, // upper left row offset (in pixels)
            0, // upper left column offset (in pixels)
            200, // height (in pixels)
            400  // width (in pixels)
        );

        // Configure HTML save options (WordArt is rendered by default)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Save the workbook as HTML with the specified options
        workbook.Save("WordArtGradient.html", htmlOptions);
    }
}