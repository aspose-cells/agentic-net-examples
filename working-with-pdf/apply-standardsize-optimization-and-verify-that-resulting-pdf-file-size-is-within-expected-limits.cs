using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfOptimizationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            for (int i = 2; i <= 101; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Product {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Configure PDF save options with Standard optimization (high print quality)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OptimizationType = PdfOptimizationType.Standard
            };

            // Define output PDF path
            string pdfPath = "StandardOptimizedOutput.pdf";

            // Save the workbook as PDF using the configured options
            workbook.Save(pdfPath, pdfOptions);

            // Verify the resulting PDF file size against an expected limit (e.g., 500 KB)
            const long expectedMaxSizeBytes = 500 * 1024; // 500 KB
            long actualSize = new FileInfo(pdfPath).Length;

            Console.WriteLine($"PDF saved to: {Path.GetFullPath(pdfPath)}");
            Console.WriteLine($"File size: {actualSize} bytes");

            if (actualSize <= expectedMaxSizeBytes)
            {
                Console.WriteLine("PDF size is within the expected limit.");
            }
            else
            {
                Console.WriteLine("PDF size exceeds the expected limit.");
            }
        }
    }
}