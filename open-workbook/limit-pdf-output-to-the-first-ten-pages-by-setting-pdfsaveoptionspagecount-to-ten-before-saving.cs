// Title: Export Only the First 10 Pages to PDF with Aspose.Cells PdfSaveOptions (C#)
// Description: This C# example loads an Excel workbook using Aspose.Cells, sets PdfSaveOptions.PageCount to 10, and saves a PDF that contains just the initial ten pages.
// Keywords: Aspose.Cells PDF page limit | PdfSaveOptions PageCount C# | export first ten pages Excel to PDF | limit PDF output Aspose.Cells | C# Aspose.Cells pagination
// Common Searches: How to save only the first 10 pages of an Excel file as PDF using Aspose.Cells in C# | Aspose.Cells PdfSaveOptions limit pages example | C# code to restrict PDF export to a set number of pages | Set page count when converting workbook to PDF with Aspose.Cells | Export a preview PDF from large workbook using Aspose.Cells
// Developer Intent: Create a PDF from an Excel workbook that includes only the first ten pages, leveraging Aspose.Cells for .NET.
// Use Cases: Generate a concise preview PDF for stakeholders without processing the entire workbook. | Produce a compliance‑ready document that contains only the opening section of a multi‑sheet report. | Automate batch conversion where each output PDF is limited to a fixed page count to reduce file size. | Provide a quick‑look PDF for mobile devices by exporting just the initial pages.
// AI Prompts: Write a C# program that loads an .xlsx file, sets PdfSaveOptions.PageCount to a variable value, and saves as PDF, handling cases where the workbook has fewer pages. | Explain how to combine PdfSaveOptions.PageCount with PdfSaveOptions.StartPage to export a custom page range using Aspose.Cells. | Show how to log the number of pages actually written when limiting PDF output with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPdfLimit
{
    // This C# example loads an Excel workbook using Aspose.Cells, sets PdfSaveOptions.PageCount to 10, and saves a PDF that contains just the initial ten pages.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (adjust the path as needed)
            Workbook workbook = new Workbook("input.xlsx");

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Limit the output to the first ten pages
            pdfOptions.PageCount = 10;

            // Save the workbook as a PDF using the configured options
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}
