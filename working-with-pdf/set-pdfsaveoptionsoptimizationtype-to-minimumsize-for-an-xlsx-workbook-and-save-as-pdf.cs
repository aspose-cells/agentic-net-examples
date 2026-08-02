using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options and configure the optimization type for smallest file size
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OptimizationType = PdfOptimizationType.MinimumSize
        };

        // Save the workbook as a PDF using the defined options
        workbook.Save("output.pdf", pdfOptions);
    }
}

// Author note: concise example demonstrating PdfSaveOptions.OptimizationType usage.