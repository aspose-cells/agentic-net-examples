// Title: Convert an Excel workbook to PDF and set the PDF creation timestamp using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook and saves it as a PDF while assigning DateTime.Now to the PDF's CreatedTime via PdfSaveOptions. | Show how to configure Aspose.Cells PdfSaveOptions to embed a custom creation date into the exported PDF. | Provide a C# example that creates a workbook, adds data, and saves it as a PDF with the current processing time recorded in the PDF metadata.
// Common Searches: how to set PDF creation date when exporting Excel to PDF with Aspose.Cells C# | Aspose.Cells PdfSaveOptions CreatedTime property usage example | C# convert workbook to PDF and include processing timestamp in metadata | set custom PDF metadata during Excel to PDF conversion Aspose.Cells | save Excel as PDF with current timestamp using Aspose.Cells .NET
// Tags: Aspose.Cells PdfSaveOptions creation timestamp | C# Excel to PDF conversion with metadata | set PDF CreatedTime Aspose.Cells | export workbook as PDF with custom timestamp | Aspose.Cells PDF metadata customization

using System;
using Aspose.Cells;

// The example creates a workbook, writes a value to cell A1, configures PdfSaveOptions with CreatedTime set to DateTime.Now, and saves the workbook as ConvertedDocument.pdf, embedding the current processing time as the PDF's creation date.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();
        // Example data
        workbook.Worksheets[0].Cells["A1"].PutValue("Converted to PDF");

        // Set PDF save options; CreatedTime reflects the processing time
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            CreatedTime = DateTime.Now
        };

        // Save the workbook as PDF using the options
        workbook.Save("ConvertedDocument.pdf", pdfOptions);
    }
}
