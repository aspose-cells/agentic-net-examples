using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some content
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("PDF/A-1a compliance demo");

        // Initialize PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // The documentation does not define a PdfACompliance property on PdfSaveOptions.
        // If such a property exists in a newer version, it would be set like this:
        // pdfOptions.PdfACompliance = PdfACompliance.PdfA1a; // <-- Missing API evidence

        // According to the available API, set the compliance level via the Compliance property
        pdfOptions.Compliance = PdfCompliance.PdfA1a;

        // Save the workbook as a PDF with the specified compliance level
        workbook.Save("output.pdf", pdfOptions);
    }
}

// Author: Aspose.Cells .NET example – sets PDF/A‑1a compliance.