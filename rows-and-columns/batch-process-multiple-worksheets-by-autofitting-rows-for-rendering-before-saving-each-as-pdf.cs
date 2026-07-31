// Title: Batch Auto‑Fit Rows and Export Each Worksheet to a One‑Page PDF with Aspose.Cells for .NET
// Description: Loads an Excel workbook, iterates through all worksheets, auto‑fits rows, sets page setup to fit the whole sheet on a single page, and saves each sheet as an individual PDF using PdfSaveOptions and SheetSet.
// Keywords: Aspose.Cells | C# auto fit rows | export worksheet to PDF | one page per sheet | PdfSaveOptions SheetSet | batch Excel to PDF | fit rows before PDF conversion
// Common Searches: auto fit rows each worksheet Aspose.Cells | save each Excel sheet as separate PDF .NET | fit entire worksheet on one PDF page | use SheetSet to render single sheet PDF | batch convert Excel workbook to PDFs
// Developer Intent: Automatically adjust row heights for every worksheet and generate a separate one‑page PDF for each sheet.
// Use Cases: Create printable PDFs where each worksheet fits on a single page for reports or invoices. | Generate individual PDFs from a multi‑sheet workbook after normalizing row heights to avoid clipping. | Automate batch processing of workbooks for archiving, sharing, or downstream workflows.
// AI Prompts: Show a C# example that auto‑fits rows on all worksheets and saves each as a one‑page PDF with Aspose.Cells. | Explain how to export PDFs without modifying the original workbook file. | Provide error‑handling patterns for missing input files and invalid sheet indices during batch PDF conversion.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace BatchAutoFitRowsToPdf
{
    // Loads an Excel workbook, iterates through all worksheets, auto‑fits rows, sets page setup to fit the whole sheet on a single page, and saves each sheet as an individual PDF using PdfSaveOptions and SheetSet.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (replace with your actual file path)
            string sourceFile = "input.xlsx";
            Workbook workbook = new Workbook(sourceFile);

            // Iterate through each worksheet in the workbook
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Auto‑fit all rows in the current worksheet
                sheet.AutoFitRows();

                // Optional: force the sheet to fit on a single page when rendered
                sheet.PageSetup.FitToPagesWide = 1;   // fit all columns on one page
                sheet.PageSetup.FitToPagesTall = 1;   // fit all rows on one page

                // Configure PDF save options for a single‑sheet PDF
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true,
                    AllColumnsInOnePagePerSheet = true,
                    // Render only the current sheet
                    SheetSet = new SheetSet(new int[] { i })
                };

                // Define the output PDF file name for the current sheet
                string outputPdf = $"Sheet{i + 1}.pdf";

                // Save the workbook (only the selected sheet) as PDF
                workbook.Save(outputPdf, pdfOptions);
            }

            Console.WriteLine("All worksheets have been auto‑fitted and saved as individual PDFs.");
        }
    }
}
