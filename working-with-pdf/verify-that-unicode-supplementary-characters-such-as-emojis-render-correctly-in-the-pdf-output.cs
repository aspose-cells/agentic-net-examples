// Title: How to verify that Unicode supplementary characters (emoji) render correctly in a PDF generated with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells PDF emoji rendering verification | check if Unicode supplementary characters appear in Aspose.Cells generated PDF | detect emoji characters in PDF output using Aspose.Cells C# | set font for emoji support when saving workbook to PDF with Aspose.Cells | read PDF bytes to confirm presence of emoji string in C#
// Tags: export PDF with emoji support Aspose.Cells | apply Segoe UI Emoji font Aspose.Cells C# | detect Unicode characters in generated PDF C# | validate PDF content for emoji Aspose.Cells | binary search for Unicode string in PDF C#

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, writes an emoji string to cell A1, applies the Segoe UI Emoji font, saves the workbook as a PDF, then reads the PDF bytes as UTF‑8 text to check whether the original emoji sequence is present, outputting a success or failure message.
class EmojiPdfVerification
{
    static void Main()
    {
        try
        {
            // Define the emoji string (Unicode supplementary characters)
            string emojiText = "Testing emojis: 🌍🚀😊";

            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put the emoji text into cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(emojiText);

            // Apply a font that supports emojis (e.g., Segoe UI Emoji)
            Style style = cell.GetStyle();
            style.Font.Name = "Segoe UI Emoji";
            cell.SetStyle(style);

            // Adjust column width to display the text fully
            sheet.AutoFitColumns();

            // Define output PDF path
            string pdfPath = Path.Combine(Environment.CurrentDirectory, "EmojiOutput.pdf");

            // Ensure the directory exists
            string pdfDir = Path.GetDirectoryName(pdfPath);
            if (!Directory.Exists(pdfDir))
                Directory.CreateDirectory(pdfDir);

            // Save the workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);

            // Verify that the generated PDF contains the original emoji string
            bool containsEmoji = false;
            if (File.Exists(pdfPath))
            {
                // Read the PDF file as binary and search for the UTF-8 bytes of the emoji text
                byte[] pdfBytes = File.ReadAllBytes(pdfPath);
                string pdfContent = System.Text.Encoding.UTF8.GetString(pdfBytes);
                containsEmoji = pdfContent.Contains(emojiText);
            }

            // Output verification result
            Console.WriteLine("PDF generated at: " + pdfPath);
            Console.WriteLine("Verification result: " + (containsEmoji ? "SUCCESS - Emoji rendered correctly." : "FAILURE - Emoji missing."));
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
