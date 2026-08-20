// Title: C# – Export Aspose.Cells Workbook to PDF with All Columns on One Page (PdfSaveOptions.AllColumnsInOnePagePerSheet)
// Description: Creates a workbook, fills the first sheet with 50 columns of sample data, configures PdfSaveOptions.OnePagePerSheet and AllColumnsInOnePagePerSheet to scale every column onto a single PDF page, and saves the result as AllColumnsOnePage.pdf.
// Keywords: Aspose.Cells PDF export | PdfSaveOptions AllColumnsInOnePagePerSheet | OnePagePerSheet C# | fit columns single PDF page | export wide worksheet to PDF .NET | C# Aspose.Cells PDF options | scale columns to page | Aspose.Cells example PDF | workbook to PDF C# | Aspose.Cells PDF save options
// Common Searches: Aspose.Cells fit all columns on one PDF page | PdfSaveOptions OnePagePerSheet C# example | How to export wide Excel sheet to single-page PDF using Aspose.Cells | AllColumnsInOnePagePerSheet usage .NET | C# code to save workbook as PDF with columns scaled | Aspose.Cells PDF options for large worksheets
// Developer Intent: Generate a PDF where each worksheet’s columns are automatically scaled to fit a single page.
// Use Cases: Produce printable reports that contain many columns but must stay on one page per sheet. | Automate conversion of wide data tables into compact PDFs for email attachments or dashboards. | Create archival PDFs of worksheets while minimizing page count for storage efficiency. | Deliver single‑page PDF snapshots of Excel dashboards for quick stakeholder review.
// AI Prompts: Show how to change the PDF page size while keeping AllColumnsInOnePagePerSheet enabled. | Provide an example of using PdfSaveOptions to fit all rows on one page instead of columns. | Explain how to apply AllColumnsInOnePagePerSheet to every worksheet in a multi‑sheet workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsPdfExport
{
    // Creates a workbook, fills the first sheet with 50 columns of sample data, configures PdfSaveOptions.OnePagePerSheet and AllColumnsInOnePagePerSheet to scale every column onto a single PDF page, and saves the result as AllColumnsOnePage.pdf.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with enough columns to demonstrate fitting all columns on one page
            for (int col = 0; col < 50; col++)
            {
                // Header row
                sheet.Cells[0, col].PutValue("Column " + (col + 1));
                // Sample data row
                sheet.Cells[1, col].PutValue("Data " + (col + 1));
            }

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure the entire sheet is rendered on a single page
            pdfOptions.OnePagePerSheet = true;

            // Fit all columns of each sheet onto one page (ignores paper width)
            pdfOptions.AllColumnsInOnePagePerSheet = true;

            // Save the workbook as PDF using the configured options
            workbook.Save("AllColumnsOnePage.pdf", pdfOptions);

            Console.WriteLine("PDF saved successfully with all columns on one page per sheet.");
        }
    }
}
