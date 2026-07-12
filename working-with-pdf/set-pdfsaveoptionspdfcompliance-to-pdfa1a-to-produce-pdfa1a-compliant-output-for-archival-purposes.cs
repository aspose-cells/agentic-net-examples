using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample content
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("PDF/A-1a compliance example");

        // Configure PDF save options to use PDF/A-1a compliance
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1a
        };

        // Save the workbook as a PDF file with the specified compliance level
        workbook.Save("output_pdfa1a.pdf", pdfOptions);
    }
}

// Author: Aspose.Cells .NET example code.