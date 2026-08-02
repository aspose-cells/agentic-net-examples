// Title: Hide rows 50‑55 and export visible data to PDF with Aspose.Cells for .NET
// Description: Loads an existing workbook, hides rows 50 through 55 on the first worksheet using Cells.HideRows, and saves the file as a PDF. The PdfSaveOptions automatically omit hidden rows, producing a document that contains only the visible content.
// Keywords: Aspose.Cells hide rows C# | Excel to PDF Aspose.Cells | PdfSaveOptions hidden rows | C# hide rows before PDF export | Aspose.Cells row visibility
// Common Searches: Aspose.Cells hide rows 50 to 55 C# | Export only visible rows to PDF using Aspose.Cells | C# hide Excel rows then save as PDF | PdfSaveOptions exclude hidden rows Aspose | How to programmatically hide rows in Aspose.Cells
// Developer Intent: Programmatically hide rows 50‑55 in a worksheet and generate a PDF that includes only the rows that remain visible.
// Use Cases: Create a printable report that excludes confidential rows by hiding them before PDF conversion. | Produce a clean PDF version of a sheet after collapsing rows used for intermediate calculations. | Deliver a dashboard PDF where specific rows (e.g., 50‑55) contain data that should not be shown to end users.
// AI Prompts: Generate C# code with Aspose.Cells that hides rows 50‑55 and saves the worksheet as a PDF, ensuring hidden rows are not rendered. | Show how to use PdfSaveOptions in Aspose.Cells to export only visible rows after calling Cells.HideRows. | Explain the behavior of hidden rows during PDF export in Aspose.Cells and how to control it with save options.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an existing workbook, hides rows 50 through 55 on the first worksheet using Cells.HideRows, and saves the file as a PDF. The PdfSaveOptions automatically omit hidden rows, producing a document that contains only the visible content.
class Program
{
    static void Main()
    {
        // Load the existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (you can change the index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Hide rows 50 to 55 (Excel rows are 1‑based, Cells API is 0‑based)
        // Start index = 49 (row 50), total rows to hide = 6 (rows 50‑55)
        worksheet.Cells.HideRows(49, 6);

        // Prepare PDF save options (default behavior skips hidden rows)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as PDF; only visible rows will appear in the output
        workbook.Save("output.pdf", pdfOptions);
    }
}
