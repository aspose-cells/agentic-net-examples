// Title: Unhide Column B, Set Width to 50 Points, and Export Worksheet to PDF with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, unhides column B (zero‑based index 1), sets its width to 50 points using Cells.UnhideColumn, and saves the first worksheet as a PDF with default PdfSaveOptions.
// Keywords: Aspose.Cells | C# | unhide column | set column width points | export worksheet to PDF | PdfSaveOptions | Excel to PDF conversion | column formatting Aspose.Cells
// Common Searches: Aspose.Cells how to unhide a column and set width before PDF export | C# set column B width to 50 points with Aspose.Cells | Export Excel worksheet to PDF after adjusting column visibility Aspose.Cells | Unhide column and define width in points using Aspose.Cells .NET | Save workbook as PDF with custom column width Aspose.Cells
// Developer Intent: Unhide column B, assign a width of 50 points, and generate a PDF of the worksheet using Aspose.Cells for .NET.
// Use Cases: Prepare a printable report where a hidden column must be visible and precisely sized before PDF generation. | Create an invoice PDF that requires column B to be displayed at a fixed width for proper alignment of data. | Generate a PDF version of a spreadsheet for distribution, ensuring hidden columns are revealed and formatted correctly.
// AI Prompts: Write C# code with Aspose.Cells to unhide column C, set its width to 30 points, and export the sheet to PDF. | Show how to customize PdfSaveOptions (e.g., page size, compression) after modifying column visibility with Aspose.Cells. | Explain the difference between setting column width in points versus characters in Aspose.Cells and provide sample code for each.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // Loads an Excel workbook, unhides column B (zero‑based index 1), sets its width to 50 points using Cells.UnhideColumn, and saves the first worksheet as a PDF with default PdfSaveOptions.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet's cells collection
            Cells cells = workbook.Worksheets[0].Cells;

            // Unhide column B (zero‑based index 1) and set its width to 50 points
            cells.UnhideColumn(1, 50);

            // Prepare PDF save options (default options are sufficient for this task)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Export the workbook (first worksheet) to PDF
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}
