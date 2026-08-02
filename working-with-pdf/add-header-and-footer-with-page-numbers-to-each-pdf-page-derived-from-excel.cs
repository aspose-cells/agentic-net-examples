using System;
using Aspose.Cells;

namespace AsposeCellsPdfHeaderFooter
{
    // Author: Aspose.Cells .NET example – adds header/footer with page numbers and saves as PDF
    class Program
    {
        static void Main()
        {
            // Load an existing Excel workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets and set header/footer
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Header: left section – "Page X of Y"
                sheet.PageSetup.SetHeader(0, "Page &P of &N");

                // Footer: center section – "Page X of Y"
                sheet.PageSetup.SetFooter(1, "Page &P of &N");
            }

            // Prepare PDF save options (default options are sufficient for header/footer)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF; each page will contain the defined header/footer
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}