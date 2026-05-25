using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfHeadingsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate some sample data with a header row
            worksheet.Cells["A1"].PutValue("Product");
            worksheet.Cells["B1"].PutValue("Price");
            worksheet.Cells["A2"].PutValue("Apple");
            worksheet.Cells["B2"].PutValue(2.5);
            worksheet.Cells["A3"].PutValue("Orange");
            worksheet.Cells["B3"].PutValue(1.8);

            // Enable printing of row and column headings on each page
            worksheet.PageSetup.PrintHeadings = true;

            // (Optional) Define the print area to include the data range
            worksheet.PageSetup.PrintArea = "A1:B3";

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with headings enabled
            workbook.Save("OutputWithHeadings.pdf", pdfOptions);
        }
    }
}