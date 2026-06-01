using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace WordArtPdfDemoApp
{
    class WordArtPdfDemo
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Add a WordArt shape with a gradient preset style (WordArtStyle7)
                // Parameters: style, text, topRow, top, leftColumn, left, height, width
                Shape wordArt = worksheet.Shapes.AddWordArt(
                    PresetWordArtStyle.WordArtStyle7,
                    "Gradient WordArt",
                    2,   // topRow
                    0,   // top (pixel offset)
                    2,   // leftColumn
                    0,   // left (pixel offset)
                    200, // height (pixels)
                    400  // width (pixels)
                );

                // Optional: customize the text effect of the WordArt
                if (wordArt.IsWordArt)
                {
                    TextEffectFormat textEffect = wordArt.TextEffect;
                    textEffect.FontBold = true;
                    textEffect.FontSize = 24;
                    textEffect.FontName = "Arial";
                }

                // Configure PDF save options (default options are sufficient)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF with the specified options
                workbook.Save("WordArtGradient.pdf", pdfOptions);
                Console.WriteLine("PDF saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}