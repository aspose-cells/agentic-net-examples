using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class VerifyUnicodeEmojiPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert text containing Unicode supplementary characters (e.g., emojis)
        worksheet.Cells["A1"].PutValue("Hello 🌍🚀");

        // Set the workbook's default font to one that supports emojis
        workbook.DefaultStyle.Font.Name = "Segoe UI Emoji";

        // Configure PDF save options for proper Unicode rendering
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Fallback font if a cell's style font cannot render the character
            DefaultFont = "Segoe UI Emoji",
            // Use Identity encoding to preserve all Unicode characters
            FontEncoding = PdfFontEncoding.Identity,
            // Try workbook's default font first
            CheckWorkbookDefaultFont = true,
            // Keep font compatibility checking enabled for substitution when needed
            CheckFontCompatibility = true
        };

        // Save the workbook as PDF
        workbook.Save("UnicodeEmoji.pdf", pdfOptions);

        Console.WriteLine("PDF saved with Unicode supplementary characters.");
    }
}