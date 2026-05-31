using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with a gradient fill (WordArtStyle7)
        // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7, // Gradient Fill - Blue, Accent 1, Reflection
            "Gradient WordArt",
            2, 0,   // upper left row and top offset
            2, 0,   // upper left column and left offset
            100,    // height
            400);   // width

        // Configure PDF save options to produce PDF/A‑2b compliant output
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.Compliance = PdfCompliance.PdfA2b; // Set PDF/A‑2b compliance

        // Save the workbook as a PDF file; gradients in WordArt are preserved automatically
        workbook.Save("WordArt_PdfA2b.pdf", pdfOptions);
    }
}