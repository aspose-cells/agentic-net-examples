// Title: How to Repeat a Header Row on Every PDF Page with Aspose.Cells for .NET (C#)
// Description: C# sample that builds a workbook, adds a header row, populates 100 data rows, configures PageSetup.PrintTitleRows to repeat the first row on each printed page, sets the print area, and saves the sheet as a multi‑page PDF using PdfSaveOptions. The generated PDF shows the header on every page for clear readability.
// Keywords: Aspose.Cells C# PDF repeat header | PrintTitleRows | page setup Aspose.Cells | export worksheet to PDF | repeat rows on PDF pages | Aspose.Cells PDF header example | C# Aspose.Cells PDF export
// Common Searches: Aspose.Cells repeat header row PDF | C# set PrintTitleRows Aspose.Cells | how to keep column headings on each PDF page using Aspose.Cells | define print area and repeat rows in PDF with Aspose.Cells | Aspose.Cells PDF export with repeating titles
// Developer Intent: Configure a worksheet so that the first row is printed as a title on every page of the generated PDF.
// Use Cases: Multi‑page PDF reports with persistent column headings | Printable invoices or statements where the header stays visible | Large data tables exported to PDF for audit or compliance documentation | Financial statements that require repeated titles across pages
// AI Prompts: Generate C# code with Aspose.Cells that repeats rows 1‑2 as titles on each PDF page and saves the workbook. | Explain the steps to set PageSetup.PrintTitleRows and PrintArea before exporting a worksheet to PDF using Aspose.Cells. | Provide troubleshooting tips when header rows do not appear on every PDF page in an Aspose.Cells export.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // C# sample that builds a workbook, adds a header row, populates 100 data rows, configures PageSetup.PrintTitleRows to repeat the first row on each printed page, sets the print area, and saves the sheet as a multi‑page PDF using PdfSaveOptions. The generated PDF shows the header on every page for clear readability.
    public class PrintTitleRowsToPdfDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate header row (row 1) and some sample data
                worksheet.Cells["A1"].PutValue("Header 1");
                worksheet.Cells["B1"].PutValue("Header 2");
                worksheet.Cells["C1"].PutValue("Header 3");

                for (int i = 2; i <= 100; i++)
                {
                    worksheet.Cells[$"A{i}"].PutValue($"Data A{i - 1}");
                    worksheet.Cells[$"B{i}"].PutValue($"Data B{i - 1}");
                    worksheet.Cells[$"C{i}"].PutValue($"Data C{i - 1}");
                }

                // Configure page setup to repeat the first row on each printed page
                PageSetup pageSetup = worksheet.PageSetup;
                pageSetup.PrintTitleRows = "$1:$1"; // repeat row 1 as title rows

                // Define the print area to include all used cells
                pageSetup.PrintArea = "A1:C100";

                // Create PDF save options (default options are sufficient for this demo)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF; the header row will repeat on each page
                workbook.Save("PrintTitleRowsDemo.pdf", pdfOptions);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
