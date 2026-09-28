// Title: Hide rows 5‑10 in an Excel worksheet and save the workbook as PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file, hides rows 5 through 10 on the first worksheet, and saves the workbook as a PDF with Aspose.Cells. | Create a reusable method that takes an input Excel path, a start row and count, hides that row range, and outputs a PDF using default PdfSaveOptions. | Refactor existing Aspose.Cells code to accept dynamic row‑range parameters, hide those rows, and then convert the workbook to PDF.
// Common Searches: Aspose.Cells hide rows 5 to 10 before saving as PDF in C# | C# hide a specific range of rows in Excel and export to PDF with Aspose.Cells | How to keep rows hidden when converting an Excel file to PDF using Aspose.Cells
// Tags: Aspose.Cells hide specific row range | Aspose.Cells PDF export preserving hidden rows | C# hide Excel rows before PDF conversion | Aspose.Cells hide rows worksheet | Excel to PDF conversion with hidden rows

using System;
using Aspose.Cells;

// Loads input.xlsx, hides rows 5‑10 on the first worksheet via Cells.HideRows, and saves the workbook as output.pdf using default PdfSaveOptions.
class Program
{
    static void Main()
    {
        // Load the existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide rows 5 through 10.
        // Cells.HideRows uses zero‑based row index and the number of rows to hide.
        // Row 5 (1‑based) => index 4, total rows to hide = 6 (5,6,7,8,9,10)
        worksheet.Cells.HideRows(4, 6);

        // Prepare PDF save options (default options are sufficient)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}
