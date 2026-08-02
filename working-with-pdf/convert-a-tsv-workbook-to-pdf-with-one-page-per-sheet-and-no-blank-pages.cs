using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // for PdfSaveOptions

// Author: Aspose.Cells .NET example – converts a TSV workbook to PDF,
// one page per sheet, and omits completely blank pages.
class Program
{
    static void Main()
    {
        // Load the TSV file into a workbook.
        // LoadOptions with LoadFormat.Tsv tells Aspose.Cells to treat the file as tab‑separated values.
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
        Workbook workbook = new Workbook("input.tsv", loadOptions);

        // Configure PDF save options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Ensure each worksheet is rendered on a single page.
            OnePagePerSheet = true,

            // Do not generate a blank page when a sheet has no printable content.
            OutputBlankPageWhenNothingToPrint = false
        };

        // Save the workbook as a PDF file using the configured options.
        workbook.Save("output.pdf", pdfOptions);
    }
}