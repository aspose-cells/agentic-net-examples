using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load the existing XLS file
        Workbook workbook = new Workbook("input.xls");

        // Create PDF save options and set the optimization type to MinimumSize
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

        // Save the workbook as a PDF using the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}