using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHeaderPdfExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for PDF export");

            // Set the left section of the page header to display the file name (&F)
            // This header will appear on every PDF page generated from the worksheet
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.PageSetup.SetHeader(0, "&F"); // 0 = left section, &F = file name without path
            }

            // Configure PDF save options (optional settings can be added here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: show the document title in the PDF viewer's title bar
                DisplayDocTitle = true
            };

            // Save the workbook as a PDF; the header defined above will be included on each page
            workbook.Save("WorkbookWithHeader.pdf", pdfOptions);
        }
    }
}