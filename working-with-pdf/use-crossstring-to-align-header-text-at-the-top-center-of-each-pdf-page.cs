using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsHeaderDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data (optional, just to have content in the sheet)
            sheet.Cells["A1"].PutValue("Sample data for PDF export");

            // Access the page setup of the worksheet
            PageSetup pageSetup = sheet.PageSetup;

            // Ensure header/footer margins align with page margins (default is true)
            pageSetup.IsHFAlignMargins = true;

            // Set the center header (section 1) with the desired text
            // This will appear at the top center of each PDF page
            pageSetup.SetHeader(1, "My Centered Header");

            // Create PDF save options (optional customizations can be added here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file; the header will be rendered on each page
            workbook.Save("HeaderCentered.pdf", pdfOptions);
        }
    }
}