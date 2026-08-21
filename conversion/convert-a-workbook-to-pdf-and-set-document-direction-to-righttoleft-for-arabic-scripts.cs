// Title: Aspose.Cells C# – Convert Workbook to PDF with Right‑to‑Left Layout for Arabic
// Description: Creates a workbook, enables the DisplayRightToLeft flag on the worksheet, adds Arabic text, configures PdfSaveOptions with an Arabic‑compatible font, and saves the file as a right‑to‑left PDF. Demonstrates proper RTL rendering for Arabic scripts in .NET.
// Keywords: Aspose.Cells | C# | .NET | PDF conversion | right to left | RTL | Arabic | DisplayRightToLeft | PdfSaveOptions | Arabic font | Middle East | Saudi Arabia | UAE
// Common Searches: Aspose.Cells set RTL when exporting to PDF | C# convert Excel to PDF with Arabic right‑to‑left layout | DisplayRightToLeft property PDF Aspose.Cells | PdfSaveOptions Arabic font missing glyphs | How to generate RTL PDF for Arabic using Aspose.Cells
// Developer Intent: Export an Excel workbook to a PDF that displays Arabic content in a right‑to‑left orientation.
// Use Cases: Generate Arabic reports or invoices with correct RTL formatting. | Create multilingual PDFs where Arabic sheets need RTL layout. | Automate batch conversion of Excel files to RTL PDFs for Middle Eastern markets.
// AI Prompts: Write C# code using Aspose.Cells to export a workbook to PDF with RTL layout and an Arabic‑compatible font. | Explain how DisplayRightToLeft and PdfSaveOptions interact to produce right‑to‑left PDFs in Aspose.Cells. | Provide a script that iterates through all worksheets, sets each to RTL, and saves each as a separate PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsArabicRtlPdf
{
    // Creates a workbook, enables the DisplayRightToLeft flag on the worksheet, adds Arabic text, configures PdfSaveOptions with an Arabic‑compatible font, and saves the file as a right‑to‑left PDF. Demonstrates proper RTL rendering for Arabic scripts in .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: Workbook constructor)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Set the worksheet to display from right to left (Arabic script direction)
            worksheet.DisplayRightToLeft = true;

            // Add some Arabic sample text
            worksheet.Cells["A1"].PutValue("مرحبا بالعالم"); // "Hello World" in Arabic

            // Configure PDF save options (optional: set a font that supports Arabic)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Use a font that contains Arabic glyphs
                DefaultFont = "Arial",
                // Ensure the default workbook font is checked for missing glyphs
                CheckWorkbookDefaultFont = true
            };

            // Save the workbook as PDF using the save options (lifecycle rule: Save with SaveOptions)
            workbook.Save("ArabicRightToLeft.pdf", pdfOptions);

            Console.WriteLine("Workbook has been saved to PDF with right‑to‑left direction.");
        }
    }
}
