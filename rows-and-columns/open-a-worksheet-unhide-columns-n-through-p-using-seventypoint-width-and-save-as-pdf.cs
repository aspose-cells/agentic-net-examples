// Title: Aspose.Cells for .NET – Unhide Columns N‑P (70 pt), Save Worksheet as PDF
// Description: Load an Excel workbook, unhide columns N through P on the first worksheet, set each column’s width to 70 points, and export the sheet to a PDF using Aspose.Cells PdfSaveOptions.
// Keywords: Aspose.Cells | C# | unhide columns | set column width points | export to PDF | PdfSaveOptions | column N to P | Excel to PDF conversion | worksheet column visibility | 70 point column width
// Common Searches: Aspose.Cells unhide columns N to P | set column width in points Aspose.Cells | export Excel sheet to PDF after changing column visibility | C# hide/unhide columns Aspose.Cells | how to set column width in points before PDF export | Aspose.Cells PdfSaveOptions example
// Developer Intent: Unhide columns N‑P, apply a 70‑point width, and generate a PDF of the worksheet.
// Use Cases: Generate printable reports where hidden columns N‑P must be visible with consistent width. | Automate financial statements that require specific column widths before PDF distribution. | Batch process multiple workbooks to reveal selected columns and produce uniform PDFs.
// AI Prompts: Write C# code with Aspose.Cells to unhide columns 13‑15, set each column width to 70 points, and save the workbook as a PDF. | Show how to hide columns again after exporting to PDF using Aspose.Cells. | Explain how to convert an Excel column letter to a zero‑based index and set its width in points with Aspose.Cells.

using Aspose.Cells;
using System;

// Load an Excel workbook, unhide columns N through P on the first worksheet, set each column’s width to 70 points, and export the sheet to a PDF using Aspose.Cells PdfSaveOptions.
class Program
{
    static void Main()
    {
        // Load an existing workbook (provide the correct path to your Excel file)
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Unhide columns N (index 13) through P (index 15) and set their width to 70 points
        // Parameters: start column index, number of columns, width in points
        worksheet.Cells.UnhideColumns(13, 3, 70);

        // Save the modified workbook as a PDF file
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        workbook.Save("output.pdf", pdfOptions);
    }
}
