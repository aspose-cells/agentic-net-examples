// Title: C# – Export Workbook to PDF with OnePagePerSheet – Each Worksheet on a New Page (Aspose.Cells)
// Description: Demonstrates how to create a workbook with multiple worksheets, fill them with data, and save the file as a PDF using Aspose.Cells PdfSaveOptions with OnePagePerSheet enabled so that every worksheet begins on its own PDF page.
// Keywords: Aspose.Cells OnePagePerSheet | C# PDF export worksheet per page | PdfSaveOptions OnePagePerSheet example | Aspose.Cells export multiple sheets to PDF | C# generate PDF with separate worksheet pages
// Common Searches: Aspose.Cells OnePagePerSheet C# | save each worksheet on a new PDF page Aspose | C# convert workbook to PDF separate sheets | PdfSaveOptions OnePagePerSheet usage | export multi‑sheet Excel to PDF Aspose.Cells
// Developer Intent: Produce a single PDF where each worksheet of a workbook is rendered on a distinct page by setting PdfSaveOptions.OnePagePerSheet to true.
// Use Cases: Financial statements with separate sheets for balance sheet, income statement, and cash flow, each printed on its own PDF page. | Departmental reports compiled into one PDF document while preserving clear page breaks between sheets. | Automated generation of printable PDFs from data‑entry worksheets, ensuring each sheet starts on a new page for readability.
// AI Prompts: Show C# code that adds three worksheets, populates them, and saves to a PDF with OnePagePerSheet enabled using Aspose.Cells. | Explain how OnePagePerSheet interacts with other PdfSaveOptions such as PageSetup, scaling, and image quality. | Provide a step‑by‑step guide to configure PdfSaveOptions for separate worksheet pages when converting an Excel workbook to PDF in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook with multiple worksheets, fill them with data, and save the file as a PDF using Aspose.Cells PdfSaveOptions with OnePagePerSheet enabled so that every worksheet begins on its own PDF page.
    public class OnePagePerSheetDemo
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("PDF generated successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (contains one default worksheet)
            Workbook workbook = new Workbook();

            // Rename the default worksheet for clarity
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";

            // Add a second worksheet
            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");

            // Fill both worksheets with enough rows to normally span multiple pages
            for (int i = 0; i < 100; i++)
            {
                sheet1.Cells[i, 0].PutValue($"Sheet1 Row {i + 1}");
                sheet2.Cells[i, 0].PutValue($"Sheet2 Row {i + 1}");
            }

            // Set PDF save options: each worksheet will be rendered on a separate page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook to PDF using the configured options
            workbook.Save("WorksheetsOnePagePerSheet.pdf", pdfOptions);
        }
    }
}
