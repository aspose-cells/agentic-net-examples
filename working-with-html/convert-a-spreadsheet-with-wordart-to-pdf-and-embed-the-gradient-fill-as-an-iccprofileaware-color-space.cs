using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class WordArtToPdfWithIccProfile
{
    static void Main()
    {
        // Path to the source Excel file (should contain WordArt or will be created)
        string inputPath = "WordArtSource.xlsx";
        // Path to the output PDF file
        string outputPdfPath = "WordArtOutput.pdf";

        Workbook workbook;

        // Load existing workbook if it exists, otherwise create a new one
        if (File.Exists(inputPath))
        {
            workbook = new Workbook(inputPath);
        }
        else
        {
            workbook = new Workbook();
            // Add a worksheet and some sample data (optional)
            Worksheet ws = workbook.Worksheets[0];
            ws.Cells["A1"].PutValue("Sample Data");
        }

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with a preset style that supports gradient fill
        // Using WordArtStyle6 (Gradient Fill - Gray) as an example
        ShapeCollection shapes = worksheet.Shapes;
        Shape wordArt = shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle6,   // preset style with gradient capability
            "Aspose.Cells",                     // text displayed in WordArt
            2, 0,                               // top row and vertical offset (pixels)
            2, 0,                               // left column and horizontal offset (pixels)
            100, 400);                          // height and width (pixels)

        // Configure the gradient fill for the WordArt
        // Set a two‑color gradient from LightGray to DarkGray, diagonal down style
        wordArt.Fill.FillType = FillType.Gradient;
        GradientFill gradientFill = wordArt.Fill.GradientFill;
        gradientFill.SetTwoColorGradient(
            Color.LightGray,                    // first gradient color
            Color.DarkGray,                     // second gradient color
            GradientStyleType.DiagonalDown,    // gradient style
            1);                                 // variant (1‑4)

        // Prepare PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Use PDF/A‑2u compliance which requires an ICC profile (color‑managed PDF)
        pdfOptions.Compliance = PdfCompliance.PdfA2u;

        // Embed standard Windows fonts to ensure proper rendering of text
        pdfOptions.EmbedStandardWindowsFonts = true;

        // Optional: set default font in case some characters lack explicit font definition
        pdfOptions.DefaultFont = "Arial";

        // Save the workbook as PDF with the specified options
        workbook.Save(outputPdfPath, pdfOptions);

        Console.WriteLine($"Workbook converted to PDF with ICC‑profile‑aware color space: {outputPdfPath}");
    }
}