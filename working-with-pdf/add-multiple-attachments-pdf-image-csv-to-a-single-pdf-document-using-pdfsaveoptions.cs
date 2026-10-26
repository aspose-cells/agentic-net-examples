// Title: How to embed multiple file attachments (PDF, image, CSV) into a PDF generated with Aspose.Cells PdfSaveOptions in C# (feature not supported)
// Common Searches: Can Aspose.Cells PdfSaveOptions add file attachments to a PDF in C#? | Embedding an image and a CSV as attachments in a PDF created by Aspose.Cells .NET | Workaround for attaching files to a PDF saved from an Aspose.Cells workbook | Why does Aspose.Cells not support PDF attachments when saving with PdfSaveOptions? | Adding multiple attachments to a PDF generated from an Excel workbook using Aspose.Cells
// Tags: Aspose.Cells PDF attachment limitation | PdfSaveOptions unsupported attachment feature | C# embed files in PDF using Aspose.Cells | generate PDF without attachments Aspose.Cells | Aspose.Cells workbook to PDF attachment workaround

using System;
using System.IO;
using Aspose.Cells;

// The example creates a simple Workbook, writes data to cell A1, configures PdfSaveOptions, and saves the workbook as ResultWithAttachments.pdf. It also notes that Aspose.Cells does not currently provide a way to embed additional files (PDF, image, CSV) as attachments in the generated PDF.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF");

            // Initialize PDF save options (no attachment support in Aspose.Cells)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            string outputPath = "ResultWithAttachments.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
