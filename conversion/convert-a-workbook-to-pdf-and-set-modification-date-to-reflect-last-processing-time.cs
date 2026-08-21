// Title: Convert an Aspose.Cells Workbook to PDF and set the PDF creation date (CreatedTime) in C#
// Description: Creates or loads a Workbook, writes a sample value, configures PdfSaveOptions.CreatedTime with DateTime.Now, and saves the workbook as a PDF file that records the generation timestamp.
// Keywords: Aspose.Cells PDF conversion | PdfSaveOptions CreatedTime | C# export Excel to PDF | set PDF creation date | Aspose.Cells metadata | Excel to PDF timestamp | .NET PDF export
// Common Searches: Aspose.Cells set PDF creation time C# | PdfSaveOptions CreatedTime example | Export Excel workbook to PDF with timestamp | How to add generation date to PDF using Aspose.Cells | C# Aspose.Cells PDF metadata settings
// Developer Intent: Export an Excel workbook to PDF while embedding the current processing time as the PDF’s creation date.
// Use Cases: Generate daily reports where each PDF shows the exact generation time for audit purposes. | Automate a web service that returns Excel data as a PDF with a timestamp for compliance tracking. | Create archival PDFs from workbooks that need a reliable creation date embedded in the file metadata.
// AI Prompts: Show how to also set the PDF ModifiedTime property to the current time using Aspose.Cells. | Provide a sample that loads an existing .xlsx file, converts it to PDF, and sets both CreatedTime and Author metadata. | Explain how to convert the DateTime to a specific timezone before assigning it to PdfSaveOptions.CreatedTime.

using System;
using Aspose.Cells;

// Creates or loads a Workbook, writes a sample value, configures PdfSaveOptions.CreatedTime with DateTime.Now, and saves the workbook as a PDF file that records the generation timestamp.
class WorkbookToPdf
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        // Add some data to demonstrate the conversion
        workbook.Worksheets[0].Cells["A1"].PutValue("Converted to PDF");

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set the creation time of the PDF to the current processing time
            CreatedTime = DateTime.Now
        };

        // Save the workbook as a PDF using the specified options
        workbook.Save("ConvertedDocument.pdf", pdfOptions);
    }
}
