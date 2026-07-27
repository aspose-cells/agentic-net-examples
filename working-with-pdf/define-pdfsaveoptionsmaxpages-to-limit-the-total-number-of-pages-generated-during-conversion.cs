using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and add some data to generate multiple pages
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        for (int i = 0; i < 500; i++)
        {
            sheet.Cells[i, 0].Value = $"Row {i + 1}";
        }

        // Initialize PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Limit the total number of pages saved.
        // The Aspose.Cells API provides the PageCount property for this purpose.
        // There is no MaxPages property in the current version.
        pdfOptions.PageCount = 5; // Save only the first 5 pages

        // Save the workbook as PDF with the page limitation applied
        workbook.Save("LimitedPagesOutput.pdf", pdfOptions);
    }
}
// Author: Aspose.Cells .NET example – limits PDF output to a maximum number of pages.