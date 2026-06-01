using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load the XLSB workbook from disk
        Workbook workbook = new Workbook("input.xlsb");

        // Configure PDF save options to prioritize minimum file size
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OptimizationType = PdfOptimizationType.MinimumSize
        };

        // Save the workbook as a PDF using the specified options
        workbook.Save("output.pdf", pdfOptions);
    }
}