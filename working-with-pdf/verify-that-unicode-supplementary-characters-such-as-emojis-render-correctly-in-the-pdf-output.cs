using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace UnicodePdfVerification
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Insert a Unicode supplementary character (emoji) into a cell
            sheet.Cells["A1"].PutValue("😀"); // Grinning face emoji

            // Configure PDF save options to handle Unicode characters correctly
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Use a font that contains emoji glyphs
                DefaultFont = "Segoe UI Emoji",

                // Try to use the workbook's default font first (helps with other Unicode chars)
                CheckWorkbookDefaultFont = true,

                // Use Identity encoding to ensure all characters are preserved
                FontEncoding = PdfFontEncoding.Identity,

                // Embed fonts for characters beyond ASCII (emoji are > 127)
                EmbedStandardWindowsFonts = true,

                // Keep font compatibility checking enabled (default) for safety
                CheckFontCompatibility = true
            };

            // Save the workbook as PDF with the specified options
            string outputPath = "UnicodeEmojiOutput.pdf";
            workbook.Save(outputPath, pdfOptions);

            // At this point, open the generated PDF (outputPath) manually or via a PDF viewer
            // to verify that the emoji renders correctly and is not displayed as a block.
        }
    }
}