using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfOptimizationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (XLSX in memory)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells PDF Optimization Example");
            sheet.Cells["A2"].PutValue("This PDF is optimized for minimum file size.");

            // Create PDF save options
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // Set the optimization type to MinimumSize (file size prioritized over quality)
            pdfSaveOptions.OptimizationType = PdfOptimizationType.MinimumSize;

            // Save the workbook as a PDF using the specified options
            workbook.Save("OptimizedOutput.pdf", pdfSaveOptions);
        }
    }
}