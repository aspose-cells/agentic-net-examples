using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Author: Aspose.Cells .NET example – PDF/A‑1a compliance check
class Program
{
    static void Main()
    {
        // Create a new workbook and add sample content
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("PDF/A‑1a compliance test");

        // Configure PDF save options to enforce PDF/A‑1a compliance
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1a
        };

        // Save the workbook as a PDF file with the specified compliance level
        string outputPath = "PdfA1aOutput.pdf";
        workbook.Save(outputPath, saveOptions);

        // Aspose.Cells generates a PDF that conforms to PDF/A‑1a when the
        // Compliance property is set to PdfCompliance.PdfA1a.
        // Additional validation can be performed with a dedicated PDF/A validator if needed.
    }
}