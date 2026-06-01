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

        // Optional: put some text in a cell
        worksheet.Cells["A1"].PutValue("Demo with WordArt");

        // Add WordArt with a gradient preset (WordArtStyle6)
        // Parameters: style, text, topRow, top offset, leftColumn, left offset, height, width
        ShapeCollection shapes = worksheet.Shapes;
        Shape wordArt = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle6,
            "Gradient WordArt",
            2, 0,    // topRow, top offset (pixels)
            2, 0,    // leftColumn, left offset (pixels)
            100,     // height (pixels)
            400);    // width (pixels)

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions
        {
            // Embed all images (including WordArt) as Base64 strings
            ExportImagesAsBase64 = true,
            // Use HTML5 to allow inline SVG elements
            HtmlVersion = HtmlVersion.Html5,
            // Produce a single HTML file with all resources embedded
            SaveAsSingleFile = true
        };

        // Save the workbook as HTML
        string outputFile = "WordArtWithSvg.html";
        workbook.Save(outputFile, htmlOptions);

        Console.WriteLine($"Workbook saved as HTML to: {Path.GetFullPath(outputFile)}");
    }
}