// Title: Convert a TSV file to a PDF/A‑2b archive‑ready document using Aspose.Cells for .NET
// Common Searches: C# code to load a TSV file with Aspose.Cells and save as PDF/A-2b | How to generate PDF/A-2b compliant PDF from tab‑separated values using Aspose.Cells | Aspose.Cells PDF/A-2b export example for TSV data in .NET | Set PDF compliance to PDF/A-2b when saving a workbook in C# Aspose.Cells | LoadOptions LoadFormat.Tsv usage with PdfSaveOptions for archival PDF
// Tags: load tsv with Aspose.Cells LoadOptions | save workbook as PDF/A-2b using PdfSaveOptions | export tab‑separated values to PDF with Aspose.Cells | C# archival PDF/A-2b generation Aspose | LoadFormat.Tsv to PDF conversion example

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads a TSV file into an Aspose.Cells Workbook via LoadFormat.Tsv, configures PdfSaveOptions with Compliance set to PdfA2b, and saves the workbook as a PDF/A‑2b compliant PDF suitable for long‑term archival.
class TsvToPdfA2bConverter
{
    static void Main()
    {
        // Path to the source TSV file
        string inputPath = "input.tsv";

        // Path for the resulting PDF/A‑2b file
        string outputPath = "output.pdf";

        // Load the TSV document using the appropriate load format
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // Configure PDF save options for PDF/A‑2b compliance
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Set the PDF compliance level to PDF/A‑2b
            Compliance = PdfCompliance.PdfA2b
        };

        // Save the workbook as a PDF/A‑2b document
        workbook.Save(outputPath, pdfOptions);
    }
}
