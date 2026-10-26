// Title: Convert a macro‑enabled XLSM workbook to PDF/A‑2a with embedded images using Aspose.Cells for .NET
// Common Searches: c# aspose.cells convert xlsm to pdfa-2a preserving all images | how to export macro enabled Excel file to PDF/A‑2a with Aspose.Cells | Aspose.Cells PDF/A‑2a compliance for workbooks containing images | save xlsm as PDF/A‑2a in .NET while keeping embedded pictures | set PdfSaveOptions compliance to PdfA2a in C# Aspose.Cells example
// Tags: xlsm to pdfa2a conversion aspnet | aspose.cells pdfsaveoptions compliance pdfa2a | embed images in pdfa2a export aspnet | macro enabled workbook pdfa2a conversion | preserve excel images during pdf export c#

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads a macro‑enabled XLSM workbook, configures PdfSaveOptions.Compliance to PdfA2a (which automatically embeds all images), and saves the workbook as a PDF/A‑2a file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the source XLSM workbook (macro-enabled Excel file)
        Workbook workbook = new Workbook("input.xlsm");

        // Configure PDF save options for PDF/A‑2a compliance
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set the PDF/A compliance level to PDF/A‑2a
            Compliance = PdfCompliance.PdfA2a,

            // Ensure that all images are embedded in the PDF (default behavior)
            // No additional setting required; Aspose.Cells embeds images automatically.
            // Optionally, you can control image resolution if needed:
            // ImageResolution = 300
        };

        // Save the workbook as a PDF/A‑2a file while preserving embedded images
        workbook.Save("output.pdf", pdfOptions);
    }
}
