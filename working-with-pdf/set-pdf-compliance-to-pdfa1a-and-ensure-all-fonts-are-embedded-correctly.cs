using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Create a new workbook and add sample content
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("PDF/A-1a compliance with embedded fonts");

        // Configure PDF save options
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Set compliance level to PDF/A-1a
            Compliance = PdfCompliance.PdfA1a,
            // Ensure standard Windows fonts are embedded (fonts are always embedded for PDF/A)
            EmbedStandardWindowsFonts = true
        };

        // Save the workbook as PDF with the specified options
        workbook.Save("output_pdfa1a.pdf", saveOptions);
    }
}

// Author: Aspose.Cells .NET example – PDF/A-1a compliance with font embedding.