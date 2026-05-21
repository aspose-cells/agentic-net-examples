using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfLimitExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing Excel workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Initialize PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Limit the output to the first 10 pages
            pdfOptions.PageCount = 10;

            // Save the workbook as a PDF with the specified page limit
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}