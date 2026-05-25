using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfMarginDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Data row 1");
            sheet.Cells["A3"].PutValue("Data row 2");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(123);
            sheet.Cells["B3"].PutValue(456);

            // Configure custom page margins (values are in centimeters)
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.TopMargin = 2.0;    // 2 cm top margin
            pageSetup.BottomMargin = 1.5; // 1.5 cm bottom margin
            pageSetup.LeftMargin = 1.0;   // 1 cm left margin
            pageSetup.RightMargin = 1.0;  // 1 cm right margin

            // Create PDF save options (inherits from PaginatedSaveOptions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: keep each sheet on a single page while respecting margins
                OnePagePerSheet = true
            };

            // Save the workbook as PDF with the custom margins
            workbook.Save("CustomMarginsOutput.pdf", pdfOptions);

            Console.WriteLine("Workbook saved to PDF with custom page margins.");
        }
    }
}