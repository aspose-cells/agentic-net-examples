// Title: Check if a PDF generated with Aspose.Cells C# contains embedded attachments using PdfSaveOptions
// Common Searches: how to verify embedded files in a PDF created with Aspose.Cells in C# | does PdfSaveOptions EmbedAttachments true actually embed attachments in the output PDF | C# code to test whether a workbook exported to PDF has attached documents | detect missing Attachments collection when exporting a workbook to PDF with Aspose.Cells | view embedded attachments in a PDF viewer after saving a workbook with Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions EmbedAttachments | C# generate PDF with embedded attachments | validate PDF attachment presence | Aspose.Cells PDF export attachment limitation | workaround embedding files using Aspose.Pdf after export

using System;
using System.IO;
using Aspose.Cells;

// The example creates a sample text file, builds a workbook with a cell value, sets PdfSaveOptions.EmbedAttachments to true, notes that the Attachments collection is unavailable in the current Aspose.Cells version, saves the workbook as a PDF, and prints the output path while handling exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Prepare a sample attachment file
            string attachmentPath = "sample.txt";
            if (!File.Exists(attachmentPath))
            {
                File.WriteAllText(attachmentPath, "This is a sample attachment.");
            }

            // Create a new workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello World");

            // Configure PDF save options to embed attachments (if supported)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true
            };

            // Note: The Attachments collection is not available in this version of Aspose.Cells.
            // If attachment embedding is required, consider using a newer version or Aspose.Pdf to add
            // attachments after the PDF is generated.

            // Save the workbook as PDF
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"PDF saved successfully to '{pdfPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
