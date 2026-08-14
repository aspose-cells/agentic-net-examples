// Title: Aspose.Cells for .NET – Set PdfSaveOptions.OptimizationType to MinimumSize and Export XLSX to PDF (C#)
// Description: Shows how to build an in‑memory Workbook, add sample data, configure PdfSaveOptions with PdfOptimizationType.MinimumSize to prioritize the smallest possible PDF file, and save the workbook as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells PDF optimization | PdfSaveOptions MinimumSize | PdfOptimizationType.MinimumSize | C# export Excel to PDF | Aspose.Cells reduce PDF size | small PDF output .NET | Excel to PDF compression | minimal PDF file size | Aspose.Cells PDF save options | optimize PDF size C#
// Common Searches: Aspose.Cells set PDF optimization to MinimumSize C# | How to create smallest PDF from Excel with Aspose.Cells | PdfSaveOptions OptimizationType MinimumSize example | Export XLSX to PDF with reduced file size Aspose.Cells | C# Aspose.Cells PDF compression settings
// Developer Intent: Export an Excel workbook to PDF while minimizing the resulting file size by using the MinimumSize optimization mode.
// Use Cases: Generate lightweight PDF reports for email distribution | Create compact PDF invoices in bulk to lower storage and bandwidth costs | Archive Excel worksheets as small PDFs for long‑term retention
// AI Prompts: Write C# code that sets PdfSaveOptions.OptimizationType to MinimumSize and saves a Workbook as PDF with Aspose.Cells. | Explain the impact of PdfOptimizationType.MinimumSize on PDF quality and file size in Aspose.Cells. | Combine MinimumSize optimization with image compression and font embedding options in PdfSaveOptions.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfOptimizationExample
{
    // Shows how to build an in‑memory Workbook, add sample data, configure PdfSaveOptions with PdfOptimizationType.MinimumSize to prioritize the smallest possible PDF file, and save the workbook as a PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some sample data
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Aspose.Cells PDF Optimization Example");
            sheet.Cells["A2"].PutValue(DateTime.Now);
            sheet.Cells["B1"].PutValue(123);
            sheet.Cells["B2"].PutValue(456);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set the optimization type to MinimumSize (file size is prioritized over print quality)
            pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

            // Save the workbook as a PDF file using the specified options
            workbook.Save("OptimizedOutput.pdf", pdfOptions);

            Console.WriteLine("PDF saved with OptimizationType = MinimumSize.");
        }
    }
}
