using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Set optimization type to Standard (high print quality)
        pdfOptions.OptimizationType = PdfOptimizationType.Standard;

        // Save the workbook as a PDF using the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}