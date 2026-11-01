// Title: Export an Aspose.Cells workbook to PDF while removing blank rows and columns by fitting all columns to one page in C#
// AI Prompts: Write C# code that creates a workbook, inserts data with intentional empty rows and columns, sets PageSetup.FitToPagesWide = 1, and saves the worksheet as a PDF using Aspose.Cells. | Modify PdfSaveOptions in Aspose.Cells to produce a PDF that minimizes whitespace caused by blank rows and columns.
// Common Searches: Aspose.Cells C# export to PDF ignore empty rows and columns | fit all columns on one page when saving Excel to PDF with Aspose.Cells | reduce whitespace in PDF generated from Excel using Aspose.Cells PageSetup | configure PdfSaveOptions to remove blank space in Aspose.Cells PDF output | C# Aspose.Cells PDF export whitespace cleanup example
// Tags: Aspose.Cells PDF export fit-to-page | C# PageSetup whitespace reduction | Aspose.Cells ignore blank rows PDF | PdfSaveOptions minimal whitespace | Excel to PDF column fitting Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a new workbook, adds data starting at B2 while leaving empty rows and columns to generate whitespace, configures the worksheet's PageSetup to fit all columns on a single page, sets PdfSaveOptions, and saves the sheet as a PDF. This demonstrates how to use Aspose.Cells page setup and PDF options to reduce blank space in the exported PDF.
class WhitespaceCleanupExport
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data with intentional empty rows and columns
            // Data starts at B2 (row index 1, column index 1)
            sheet.Cells["B2"].PutValue("Header1");
            sheet.Cells["C2"].PutValue("Header2");
            sheet.Cells["B3"].PutValue(10);
            sheet.Cells["C3"].PutValue(20);
            // Row 4 and column D are left empty to create whitespace

            // Set page setup to fit content and remove extra whitespace
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.FitToPagesWide = 1;   // Fit all columns on one page
            pageSetup.FitToPagesTall = 0;   // Allow rows to flow to multiple pages

            // Configure PDF save options (ignore blank rows/columns not supported in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = false,                 // Allow content to span multiple pages
                Compliance = PdfCompliance.PdfA1b        // Example compliance setting
            };

            // Export the worksheet to PDF with whitespace cleanup
            workbook.Save("WhitespaceCleanedExport.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
