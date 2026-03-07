using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Rendering;

class ExcelToPdfA1a
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Desired output PDF file (PDF/A‑1a compliant)
        string outputPath = "output.pdf";

        // Load options – default settings (auto‑detect format)
        LoadOptions loadOptions = new LoadOptions();

        // PDF save options – set compliance to PDF/A‑1a
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.Compliance = PdfCompliance.PdfA1a;

        // Convert the Excel file to PDF using the utility method with options
        ConversionUtility.Convert(sourcePath, loadOptions, outputPath, pdfOptions);

        Console.WriteLine("Conversion to PDF/A‑1a completed successfully.");
    }
}