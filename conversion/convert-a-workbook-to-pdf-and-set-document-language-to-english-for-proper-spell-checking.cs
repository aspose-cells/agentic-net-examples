using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample text
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("This is a sample text for spell checking.");

        // Optional: set the built‑in document language property
        workbook.BuiltInDocumentProperties.Language = "en-US";

        // Configure PDF save options to use English as the default edit language
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.DefaultEditLanguage = DefaultEditLanguage.English;

        // Save the workbook as PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}