using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering;

class TsvToPdfA2b
{
    static void Main()
    {
        // Path to the source TSV file
        string sourceTsv = "input.tsv";

        // Desired output PDF file path
        string outputPdf = "output.pdf";

        // Load options specifying that the source format is TSV
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);

        // PDF save options with PDF/A‑2b compliance
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2b
        };

        // Convert the TSV file to PDF using the provided ConversionUtility method
        ConversionUtility.Convert(sourceTsv, loadOptions, outputPdf, pdfOptions);

        Console.WriteLine("TSV file has been successfully converted to PDF/A‑2b compliant PDF.");
    }
}