using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;   // For PdfCompliance enum
using Aspose.Cells.Saving;      // For PdfSaveOptions

// Author: Aspose.Cells .NET example – TSV to PDF/A‑2b conversion
class TsvToPdfA2b
{
    static void Main()
    {
        // Load the TSV file. Specify the load format explicitly.
        var loadOptions = new LoadOptions(LoadFormat.Tsv);
        var workbook = new Workbook("input.tsv", loadOptions);

        // Configure PDF save options for PDF/A‑2b compliance.
        var pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2b   // ISO 19005‑2 archival compliance
        };

        // Save the workbook as a PDF/A‑2b file.
        workbook.Save("output.pdf", pdfOptions);
    }
}