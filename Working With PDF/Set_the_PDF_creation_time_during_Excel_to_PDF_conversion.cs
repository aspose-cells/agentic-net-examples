using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfCreationTimeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory Excel file)
            Workbook workbook = new Workbook();

            // Add some sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("PDF with custom creation time");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set the desired creation time for the generated PDF
                CreatedTime = new DateTime(2023, 5, 15, 10, 30, 0),

                // (Optional) Set compliance level, e.g., PDF/A-1b
                Compliance = PdfCompliance.PdfA1b
            };

            // Save the workbook as a PDF using the configured options
            string outputPath = "WorkbookWithCreatedTime.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved to '{outputPath}' with CreatedTime = {pdfOptions.CreatedTime}");
        }
    }
}