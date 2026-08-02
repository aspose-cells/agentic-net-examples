using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all pivot tables from every worksheet in the workbook
        workbook.Worksheets.ClearPivottables();

        // Optional: configure PDF save options (e.g., ignore blank pages)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PrintingPageType = PrintingPageType.IgnoreBlank
        };

        // Save the cleaned workbook as a PDF file
        workbook.Save("clean_report.pdf", pdfOptions);
    }
}