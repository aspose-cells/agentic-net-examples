using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Author: Aspose.Cells .NET example – PDF/A‑1a compliance

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF/A‑1a compliant document");

        // Set PDF save options with PDF/A‑1a compliance
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1a
        };

        // Save the workbook as a PDF file using the specified options
        workbook.Save("output_pdfa1a.pdf", pdfOptions);
    }
}