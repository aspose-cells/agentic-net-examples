using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfOptimization
{
    // Author: Aspose.Cells .NET example – MinimumSize optimization with font subsetting
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells PDF Optimization Example");
            sheet.Cells["A2"].PutValue("This PDF uses MinimumSize optimization and font subsetting.");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Reduce file size by selecting the MinimumSize optimization mode
                OptimizationType = PdfOptimizationType.MinimumSize,

                // Enable font subsetting (embed only the glyphs that are used)
                IsFontSubstitutionCharGranularity = true
            };

            // Save the workbook as a PDF with the specified options
            workbook.Save("OptimizedDocument.pdf", pdfOptions);
        }
    }
}