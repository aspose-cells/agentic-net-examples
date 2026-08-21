// Title: Include Row and Column Headings on Every PDF Page with Aspose.Cells (C#)
// Description: Demonstrates how to enable Excel row and column headings for each page of a PDF generated with Aspose.Cells. The example creates a workbook, adds a header row, sets PrintHeadings and PrintTitleRows, applies default PdfSaveOptions, and saves the file as a multi‑page PDF with headings visible on every page.
// Keywords: Aspose.Cells PDF headings | PrintHeadings C# | repeat header row PDF Aspose | PdfSaveOptions row column headings | Excel to PDF with headings | Aspose.Cells page setup PDF | C# export worksheet to PDF
// Common Searches: how to print row and column headings in PDF using Aspose.Cells | Aspose.Cells repeat header row on each PDF page C# | enable PrintHeadings for PDF export .NET | Aspose.Cells PdfSaveOptions include headings | C# export large worksheet to PDF with headings
// Developer Intent: The developer wants every page of the PDF produced from an Excel worksheet to display the worksheet’s row and column headings for easy reference.
// Use Cases: Create multi‑page reports where each page shows Excel’s row/column labels for quick navigation. | Generate printable invoices or catalogs that retain column headers on every PDF page. | Export large data tables to PDF while keeping the first row as a repeating title across all pages.
// AI Prompts: Show how to change the font style of printed row and column headings when saving to PDF with Aspose.Cells. | Provide code to add page numbers together with row/column headings in the PDF output using PdfSaveOptions. | Explain how to disable PrintHeadings but still repeat the header row on each PDF page.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Demonstrates how to enable Excel row and column headings for each page of a PDF generated with Aspose.Cells. The example creates a workbook, adds a header row, sets PrintHeadings and PrintTitleRows, applies default PdfSaveOptions, and saves the file as a multi‑page PDF with headings visible on every page.
class PdfWithHeadings
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a header row and some sample data
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Price");
        worksheet.Cells["C1"].PutValue("Quantity");

        for (int i = 2; i <= 100; i++)
        {
            worksheet.Cells[$"A{i}"].PutValue($"Item {i - 1}");
            worksheet.Cells[$"B{i}"].PutValue(i * 1.5);
            worksheet.Cells[$"C{i}"].PutValue(i * 10);
        }

        // Enable printing of row and column headings on each page
        worksheet.PageSetup.PrintHeadings = true;

        // Repeat the header row on every printed page for better reference
        worksheet.PageSetup.PrintTitleRows = "$1:$1";

        // Create PDF save options (default settings are sufficient)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as a PDF file with headings included
        workbook.Save("ReportWithHeadings.pdf", pdfOptions);
    }
}
