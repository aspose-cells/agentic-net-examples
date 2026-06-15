using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfOptimizationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (can be based on an existing XLS file if needed)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["B1"].PutValue("More Data");
            sheet.Cells["B2"].PutValue(456);

            // Create PDF save options and set the optimization type to Standard (high print quality)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OptimizationType = PdfOptimizationType.Standard
            };

            // Save the workbook as a PDF file using the specified options
            workbook.Save("OptimizedStandard.pdf", pdfOptions);

            Console.WriteLine("PDF saved with OptimizationType = Standard.");
        }
    }
}