using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

class EmbedJavaScriptInPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and add some content
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("This PDF will open a URL on load");

            // Define output PDF path
            string outputPdfPath = "output.pdf";

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: Aspose.Cells for .NET does not provide a direct JavaScript property for PDF.
            // If JavaScript embedding is required, consider using Aspose.Pdf after conversion.

            // Save the workbook as PDF
            workbook.Save(outputPdfPath, pdfOptions);

            Console.WriteLine($"PDF saved to '{outputPdfPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}