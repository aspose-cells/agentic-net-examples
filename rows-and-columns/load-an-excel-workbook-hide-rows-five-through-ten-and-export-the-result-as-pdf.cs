// Title: Hide Rows 5‑10 in an Excel Workbook and Export to PDF using Aspose.Cells for .NET (C#)
// Description: Load an Excel file with Aspose.Cells, hide rows 5‑10 on the first worksheet via HideRows, then save the workbook as a PDF using PdfSaveOptions.
// Keywords: Aspose.Cells | C# hide rows | HideRows method | Excel to PDF conversion | PdfSaveOptions | Aspose.Cells PDF export | hide rows before PDF
// Common Searches: Aspose.Cells hide rows 5 to 10 C# | export Excel to PDF after hiding rows Aspose | C# code to hide specific rows and save workbook as PDF
// Developer Intent: Hide rows 5‑10 in an Excel workbook and generate a PDF that excludes those rows using Aspose.Cells for .NET.
// Use Cases: Create a printable PDF report that omits internal calculation rows. | Produce a client‑ready spreadsheet PDF while keeping confidential rows hidden. | Generate a clean summary PDF where grouping or header rows are excluded.
// AI Prompts: Show C# code with Aspose.Cells that hides rows 5‑10 and saves the workbook as a PDF with custom page margins. | Provide an example that conditionally hides rows based on cell values before exporting to PDF using Aspose.Cells. | Explain how to retain cell formatting while hiding rows and converting the worksheet to PDF with Aspose.Cells.

using System;
using Aspose.Cells;

// Load an Excel file with Aspose.Cells, hide rows 5‑10 on the first worksheet via HideRows, then save the workbook as a PDF using PdfSaveOptions.
class Program
{
    static void Main()
    {
        // Load the existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Hide rows 5 through 10 (zero‑based index: start at 4, hide 6 rows)
        workbook.Worksheets[0].Cells.HideRows(4, 6);

        // Prepare PDF save options (default options are sufficient)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Export the workbook to PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}
