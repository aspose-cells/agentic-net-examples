using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using System.Drawing;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add a WordArt shape that uses a preset gradient style (WordArtStyle7)
        // Parameters: style, text, topRow, top, leftColumn, left, height, width
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,
            "Gradient WordArt",
            2,          // topRow
            10,         // top (pixels)
            2,          // leftColumn
            10,         // left (pixels)
            100,        // height (pixels)
            400);       // width (pixels)

        // Give the shape a recognizable name so we can target it with CSS
        wordArt.Name = "gradientWordArt";

        // Optional: tweak the text appearance
        if (wordArt.IsWordArt)
        {
            var textEffect = wordArt.TextEffect;
            textEffect.FontSize = 24;
            textEffect.FontBold = true;
            textEffect.FontName = "Arial";
        }

        // Configure HTML export options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Add custom CSS that will be embedded in the generated HTML.
            // The CSS targets the shape by its id (the shape name becomes the element id).
            CssStyles = "#gradientWordArt { border: 2px solid #ff6600; opacity: 0.9; }",
            // Export everything into a single HTML file.
            SaveAsSingleFile = true
        };

        // Save the workbook as HTML, preserving the gradient fill and applying the custom CSS.
        workbook.Save("WordArtGradient.html", htmlOptions);
    }
}