using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class ConvertXlsToPdf
{
    static void Main()
    {
        // Source XLS file path
        string sourcePath = "input.xls";

        // Destination PDF file path
        string outputPath = "output.pdf";

        // Load the workbook from the XLS file
        Workbook workbook = new Workbook(sourcePath);

        // Configure PDF save options with MinimumSize optimization
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as PDF while preserving worksheet colors
        workbook.Save(outputPath, pdfOptions);

        Console.WriteLine("Workbook successfully converted to PDF with MinimumSize optimization.");
    }
}