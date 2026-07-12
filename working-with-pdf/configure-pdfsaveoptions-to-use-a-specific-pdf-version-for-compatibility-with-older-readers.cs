using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // PdfCompliance enum resides here

// Author: Aspose.Cells .NET example – configure PDF version compatibility
class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample text for PDF export");

        // Configure PDF save options to target PDF 1.4 for older readers
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set compliance level – Pdf14 corresponds to PDF 1.4
            Compliance = PdfCompliance.Pdf14
        };

        // Save the workbook as PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}