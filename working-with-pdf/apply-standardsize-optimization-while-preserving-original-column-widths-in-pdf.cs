using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfOptimizationDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Quantity");
            worksheet.Cells["A2"].PutValue("Apples");
            worksheet.Cells["B2"].PutValue(150);
            worksheet.Cells["A3"].PutValue("Bananas");
            worksheet.Cells["B3"].PutValue(200);
            worksheet.Cells["A4"].PutValue("Cherries");
            worksheet.Cells["B4"].PutValue(75);

            // Manually set column widths to preserve them in the PDF
            // Width is specified in characters (default unit)
            worksheet.Cells.SetColumnWidth(0, 25); // Column A
            worksheet.Cells.SetColumnWidth(1, 15); // Column B

            // Create PDF save options and set the optimization type to Standard
            // Standard corresponds to high print quality and keeps original layout
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                OptimizationType = PdfOptimizationType.Standard
            };

            // Save the workbook as a PDF while preserving the column widths
            workbook.Save("PreservedColumnsStandardOptimization.pdf", pdfSaveOptions);

            Console.WriteLine("PDF saved with Standard optimization and original column widths preserved.");
        }
    }
}