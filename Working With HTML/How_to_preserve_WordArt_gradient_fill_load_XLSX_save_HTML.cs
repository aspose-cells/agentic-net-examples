using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class PreserveWordArtGradient
{
    static void Main()
    {
        // Load the source XLSX file that contains WordArt with a gradient fill
        Workbook workbook = new Workbook("WordArtGradient.xlsx");

        // Configure HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Use HTML5 for better CSS support (optional but recommended)
        htmlOptions.HtmlVersion = HtmlVersion.Html5;

        // Export shapes (including WordArt) as images encoded in Base64.
        // This preserves complex fills such as gradients when converting to HTML.
        htmlOptions.ExportImagesAsBase64 = true;

        // Save the workbook as an HTML file while preserving the WordArt gradient fill
        workbook.Save("WordArtGradient.html", htmlOptions);
    }
}