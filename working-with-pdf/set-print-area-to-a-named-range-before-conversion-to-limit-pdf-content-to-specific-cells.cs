// Title: Set Print Area and Export Selected Cells to PDF with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to define a worksheet's print area (or named range) using PageSetup.PrintArea, configure PdfSaveOptions, and save the workbook so that only the specified cells appear in the generated PDF.
// Keywords: Aspose.Cells print area | C# export PDF selected range | PageSetup.PrintArea | PdfSaveOptions | limit PDF output cells | named range PDF conversion | Aspose.Cells .NET PDF | worksheet print area PDF | Aspose.Cells PDF save options | export specific cells to PDF
// Common Searches: Aspose.Cells set print area before PDF export | C# export only part of worksheet to PDF | How to use PageSetup.PrintArea with Aspose.Cells | PdfSaveOptions respect print area Aspose.Cells | Export named range to PDF using Aspose.Cells .NET
// Developer Intent: The developer needs to restrict the PDF output to a defined print area or named range so that only those cells are included when the workbook is saved as a PDF.
// Use Cases: Create a concise PDF report that includes only the header and the first few data rows. | Generate an invoice PDF that contains just the billing details while omitting auxiliary worksheets. | Produce a PDF snapshot of a chart’s data range by limiting the export to the chart’s source cells. | Share a printable summary of a large spreadsheet by exporting only a specific section.
// AI Prompts: Show C# code that sets a named range as the print area and saves the workbook to PDF with Aspose.Cells. | Explain how PdfSaveOptions can be configured to ensure the defined print area is honored during PDF conversion. | Provide a step‑by‑step guide to verify that the resulting PDF contains only the cells specified in the print area. | Suggest ways to programmatically create and apply a named range as the print area before exporting to PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsPrintAreaToPdf
{
    // Demonstrates how to define a worksheet's print area (or named range) using PageSetup.PrintArea, configure PdfSaveOptions, and save the workbook so that only the specified cells appear in the generated PDF.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Name");
                sheet.Cells["B1"].PutValue("Score");
                sheet.Cells["A2"].PutValue("Alice");
                sheet.Cells["B2"].PutValue(85);
                sheet.Cells["A3"].PutValue("Bob");
                sheet.Cells["B3"].PutValue(92);
                sheet.Cells["A4"].PutValue("Charlie");
                sheet.Cells["B4"].PutValue(78);

                // Define the print area directly (A1:B3)
                sheet.PageSetup.PrintArea = "A1:B3";

                // Configure PDF save options (optional)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Ensure the print area is respected
                    PrintingPageType = PrintingPageType.Default
                };

                // Save the workbook as PDF; only the defined print area will appear in the file
                workbook.Save("ExportedArea.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
