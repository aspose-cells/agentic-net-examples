// Title: C# – Export Aspose.Cells Worksheet to PDF with Embedded Fonts for Unicode Text
// Description: Creates a workbook, adds English, Chinese, and Cyrillic strings, and saves it as a PDF using PdfSaveOptions configured to embed TrueType fonts, set Arial as the fallback, enable workbook‑default font checking, and apply Identity encoding. The resulting PDF preserves the exact appearance of all characters on any device.
// Keywords: Aspose.Cells PDF export C# | embed fonts PDF Aspose.Cells | Unicode PDF Aspose.Cells | PdfSaveOptions EmbedStandardWindowsFonts | Identity font encoding Aspose.Cells | default font PDF export | multilingual PDF generation | C# Excel to PDF with fonts embedded
// Common Searches: how to embed TrueType fonts when saving Excel to PDF with Aspose.Cells | Aspose.Cells preserve Chinese characters in PDF | set default font for PDF export Aspose.Cells .NET | enable Identity encoding for PDF in Aspose.Cells | C# export worksheet to PDF with embedded fonts
// Developer Intent: Generate a PDF from a worksheet where all fonts are embedded so the text renders correctly on any platform.
// Use Cases: Produce multilingual reports (e.g., English, Chinese, Cyrillic) that retain exact typography when shared as PDFs. | Create printable invoices or contracts that must display consistently across Windows, macOS, and Linux devices. | Distribute marketing brochures worldwide without risking missing characters due to absent system fonts.
// AI Prompts: Show how to embed a custom TrueType font from a specific folder when exporting a worksheet to PDF with Aspose.Cells. | Give an example of configuring PdfSaveOptions to embed all fonts and disable substitution for Arabic text. | Explain methods to verify that fonts are correctly embedded in the generated PDF using Aspose.PDF or third‑party tools.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // Creates a workbook, adds English, Chinese, and Cyrillic strings, and saves it as a PDF using PdfSaveOptions configured to embed TrueType fonts, set Arial as the fallback, enable workbook‑default font checking, and apply Identity encoding. The resulting PDF preserves the exact appearance of all characters on any device.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Add sample data that includes Unicode characters
            worksheet.Cells["A1"].PutValue("Hello, World!");
            worksheet.Cells["A2"].PutValue("中文字符测试"); // Chinese characters
            worksheet.Cells["A3"].PutValue("Привет мир"); // Cyrillic characters

            // Configure PDF save options to embed fonts and preserve text appearance
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure TrueType fonts are embedded (affects ASCII range)
                EmbedStandardWindowsFonts = true,

                // Set a default font to use when the original font is unavailable
                DefaultFont = "Arial",

                // Try to use the workbook's default font for Unicode characters
                CheckWorkbookDefaultFont = true,

                // Use Identity encoding for all embedded fonts
                FontEncoding = PdfFontEncoding.Identity
            };

            // Save the workbook as a PDF file with the specified options
            workbook.Save("ExportedWorksheet.pdf", pdfOptions);

            Console.WriteLine("Worksheet exported to PDF with embedded fonts successfully.");
        }
    }
}
