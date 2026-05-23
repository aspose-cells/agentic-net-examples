using System;
using Aspose.Cells;

namespace AsposeCellsPdfRtlDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Enable right‑to‑left display for Arabic scripts
            sheet.DisplayRightToLeft = true;

            // Add sample Arabic text (optional, just for demonstration)
            sheet.Cells["A1"].PutValue("مرحبا بالعالم"); // "Hello World" in Arabic

            // Configure PDF save options to handle Unicode characters correctly
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            pdfOptions.CheckWorkbookDefaultFont = true;   // Use workbook's default font if needed
            pdfOptions.DefaultFont = "Arial";             // A font that supports Arabic characters

            // Save the workbook as PDF using the configured options
            workbook.Save("ArabicRightToLeft.pdf", pdfOptions);

            Console.WriteLine("Workbook successfully saved as PDF with right‑to‑left direction.");
        }
    }
}