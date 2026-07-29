// Title: Unhide Columns N‑P (70 pt width) and Export Worksheet to PDF with Aspose.Cells for .NET (C#)
// Description: C# code that loads an Excel workbook using Aspose.Cells, unhides columns N through P, sets each column's width to 70 points, and saves the worksheet as a PDF via PdfSaveOptions.
// Keywords: Aspose.Cells C# unhide columns | set column width points Aspose | export Excel to PDF .NET | unhide multiple columns Aspose.Cells | PdfSaveOptions configuration | column visibility Excel PDF
// Common Searches: Aspose.Cells how to unhide specific columns and set width | C# export Excel worksheet to PDF after adjusting column visibility | unhide columns N to P and save as PDF using Aspose.Cells
// Developer Intent: Make columns N‑P visible with a 70‑point width and generate a PDF version of the worksheet.
// Use Cases: Prepare a printable PDF report by revealing hidden columns before export. | Ensure invoice layouts render correctly in PDF after adjusting column widths. | Automate workbook cleanup: unhide hidden columns and produce a PDF for distribution.
// AI Prompts: Write C# code with Aspose.Cells to unhide columns 13‑15, set each width to 70 points, and save the sheet as a PDF. | Show how to customize PdfSaveOptions (orientation, scaling, image quality) when exporting after column adjustments. | Explain the parameters of Worksheet.Cells.UnhideColumns for setting column visibility and width in points.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// C# code that loads an Excel workbook using Aspose.Cells, unhides columns N through P, sets each column's width to 70 points, and saves the worksheet as a PDF via PdfSaveOptions.
class UnhideColumnsAndSavePdf
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Unhide columns N (13) through P (15) and set their width to 70 points
        // Parameters: start column index, number of columns, width in points
        worksheet.Cells.UnhideColumns(13, 3, 70);

        // Prepare PDF save options (optional settings can be added here)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the modified workbook as PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}
