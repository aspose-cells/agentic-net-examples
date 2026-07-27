using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfConversion
{
    // Author: Aspose.Cells .NET example – converts XLS to PDF with MinimumSize optimization while keeping colors.
    class Program
    {
        static void Main()
        {
            // Load the source XLS workbook
            Workbook workbook = new Workbook("input.xls");

            // Configure PDF save options
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                // Use MinimumSize optimization to reduce file size
                OptimizationType = PdfOptimizationType.MinimumSize
                // Worksheet colors are preserved by default; no additional setting required
            };

            // Save the workbook as a PDF with the specified options
            workbook.Save("output.pdf", pdfSaveOptions);
        }
    }
}