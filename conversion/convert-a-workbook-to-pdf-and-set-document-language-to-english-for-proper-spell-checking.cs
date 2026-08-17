// Title: Convert Excel Workbook to PDF with English Language Setting using Aspose.Cells for .NET
// Description: Demonstrates how to create an Excel workbook, add English text, configure PdfSaveOptions with DefaultEditLanguage = English, and save the workbook as a PDF so that spell‑checking tools recognize the document as English.
// Keywords: Aspose.Cells PDF conversion | C# PdfSaveOptions | DefaultEditLanguage English | Excel to PDF spell checking | set PDF language Aspose.Cells | Aspose.Cells workbook to PDF | language metadata PDF
// Common Searches: Aspose.Cells set PDF language C# | PdfSaveOptions DefaultEditLanguage example | convert Excel to PDF with English language | spell check enabled PDF from Aspose.Cells | how to set document language when saving PDF with Aspose
// Developer Intent: Generate a PDF from an Excel workbook and specify English as the default edit language to enable correct spell checking.
// Use Cases: Produce PDFs for English documents that need accurate spell‑checking in PDF viewers. | Automate report generation where the output PDF must be identified as English for accessibility tools. | Batch‑process multiple workbooks, ensuring each PDF carries the English language metadata.
// AI Prompts: Write C# code with Aspose.Cells to convert an Excel file to PDF and set DefaultEditLanguage to English. | Show how to apply PdfSaveOptions.DefaultEditLanguage = DefaultEditLanguage.English for spell‑checking support. | Modify the example to accept input and output file paths as parameters while preserving the English language setting.

using System;
using Aspose.Cells;

// Demonstrates how to create an Excel workbook, add English text, configure PdfSaveOptions with DefaultEditLanguage = English, and save the workbook as a PDF so that spell‑checking tools recognize the document as English.
class ConvertWorkbookToPdfWithEnglishLanguage
{
    static void Main()
    {
        // Create a new workbook and access the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample English text
        worksheet.Cells["A1"].PutValue("This is a sample text for spell checking.");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set the default edit language to English so that spell checking works correctly
        pdfOptions.DefaultEditLanguage = DefaultEditLanguage.English;

        // Save the workbook as a PDF file using the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}
