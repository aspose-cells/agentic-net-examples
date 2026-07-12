using System;
using Aspose.Cells;

namespace AsposeCellsPdfExample
{
    // Author: Aspose.Cells .NET example – sets IgnoreError to true for PDF conversion
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in‑memory)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data (optional – demonstrates rendering)
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(123);
            sheet.Cells["A3"].PutValue(DateTime.Now);

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Enable ignoring rendering errors (e.g., shape, chart, image issues)
            pdfOptions.IgnoreError = true;

            // Save the workbook as PDF using the configured options
            workbook.Save("OutputWithIgnoreError.pdf", pdfOptions);
        }
    }
}