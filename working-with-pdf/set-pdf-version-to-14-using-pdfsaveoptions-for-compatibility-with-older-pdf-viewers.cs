using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF 1.4 Compatibility Demo");

        // Set PDF save options to PDF 1.4 (compatible with older PDF viewers)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Pdf14 enum value corresponds to PDF version 1.4
            Compliance = PdfCompliance.Pdf14
        };

        // Save the workbook as PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}

// Author: Example demonstrating PDF 1.4 compliance with Aspose.Cells for .NET.