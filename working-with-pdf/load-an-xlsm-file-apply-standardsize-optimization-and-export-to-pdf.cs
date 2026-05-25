using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ConvertXlsmToPdfWithStandardOptimization
{
    static void Main()
    {
        // Path to the source XLSM file (must exist)
        string sourcePath = "input.xlsm";

        // Path for the resulting PDF file
        string outputPath = "output.pdf";

        // Load the XLSM workbook
        Workbook workbook = new Workbook(sourcePath);

        // Create PDF save options and set the optimization type to Standard (high print quality)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OptimizationType = PdfOptimizationType.Standard;

        // Save the workbook as a PDF using the specified options
        workbook.Save(outputPath, pdfOptions);

        Console.WriteLine($"Workbook '{sourcePath}' has been saved as PDF with Standard optimization to '{outputPath}'.");
    }
}