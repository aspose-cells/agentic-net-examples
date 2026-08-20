// Title: C# – Repeat Header Row on Each PDF Page with Aspose.Cells PrintTitleRows
// Description: Shows how to create a workbook, set a print area, designate the first row as a print title (PrintTitleRows = "$1:$1"), optionally fit the sheet to one page width, and save as PDF so the header repeats on every page.
// Keywords: Aspose.Cells C# PDF header repeat | PrintTitleRows Aspose.Cells | repeat rows on PDF | Aspose.Cells PageSetup | FitToPagesWide PDF export | export Excel to PDF with repeated header | Aspose.Cells .NET PDF export
// Common Searches: Aspose.Cells repeat header row PDF | PrintTitleRows C# Aspose.Cells | How to repeat rows on each PDF page Aspose.Cells | Fit worksheet width PDF Aspose.Cells | Set print titles for PDF export .NET
// Developer Intent: Add a repeating header row to every page of a PDF generated from an Excel worksheet using Aspose.Cells.
// Use Cases: Multi‑page PDF reports where column headings stay visible on each page. | Printable invoices that keep the first row as a static header across all pages. | Large data tables exported to PDF with the top row repeated for readability.
// AI Prompts: Show me how to set PrintTitleRows in Aspose.Cells to repeat a header row when saving to PDF. | Provide an example of using FitToPagesWide together with PrintTitleRows to produce a single‑page‑width PDF with repeated headers. | Explain how to configure PageSetup to repeat multiple rows as titles in a PDF export using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPrintTitleDemo
{
    // Shows how to create a workbook, set a print area, designate the first row as a print title (PrintTitleRows = "$1:$1"), optionally fit the sheet to one page width, and save as PDF so the header repeats on every page.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with a header row
            worksheet.Cells["A1"].PutValue("Header");
            for (int row = 2; row <= 100; row++)
            {
                worksheet.Cells[$"A{row}"].PutValue($"Data {row - 1}");
                worksheet.Cells[$"B{row}"].PutValue(row * 10);
            }

            // Set the print area to include all populated cells
            worksheet.PageSetup.PrintArea = "A1:B100";

            // Repeat the first row on each printed page
            worksheet.PageSetup.PrintTitleRows = "$1:$1";

            // Optional: fit the worksheet to a single page width
            worksheet.PageSetup.FitToPagesWide = 1;
            worksheet.PageSetup.FitToPagesTall = 0; // 0 means unlimited pages tall

            // Create PDF save options (no special options needed for titles)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF; the header row will repeat on each page
            workbook.Save("PrintTitleRowsDemo.pdf", pdfOptions);
        }
    }
}
