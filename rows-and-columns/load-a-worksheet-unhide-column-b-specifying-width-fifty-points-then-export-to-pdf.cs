// Title: C# – Unhide Column B, Set Width to 50 pts, and Save Worksheet as PDF using Aspose.Cells
// Description: Loads an Excel workbook, unhides column B (index 1), sets its width to 50 points via Cells.UnhideColumn, and saves the file as a PDF with PdfSaveOptions.
// Keywords: Aspose.Cells | C# | unhide column | set column width | points | PdfSaveOptions | export to PDF | Excel to PDF | worksheet formatting | Aspose.Cells for .NET
// Common Searches: Aspose.Cells unhide column B C# | set column width points Aspose.Cells | export Excel to PDF Aspose.Cells C# | make hidden column visible before PDF conversion Aspose.Cells | C# adjust column width and save as PDF with Aspose.Cells
// Developer Intent: Reveal column B, assign a 50‑point width, and generate a PDF from the worksheet.
// Use Cases: Create printable PDF reports where hidden columns must appear with a fixed width for layout consistency. | Generate invoices in PDF where column B holds monetary values that need a precise 50‑point width. | Automate data export pipelines that unhide specific columns, set their dimensions, and deliver the result as a PDF.
// AI Prompts: Write C# code with Aspose.Cells to unhide column index 2, set its width to 70 points, and save the workbook as a PDF. | Provide a reusable method that takes a worksheet, column index, width (points), and PDF path, then unhides the column, adjusts its width, and exports to PDF using Aspose.Cells. | Explain how to change column visibility and width in points before converting an Excel file to PDF with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, unhides column B (index 1), sets its width to 50 points via Cells.UnhideColumn, and saves the file as a PDF with PdfSaveOptions.
class UnhideColumnAndExportPdf
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Unhide column B (zero‑based index 1) and set its width to 50 points
        worksheet.Cells.UnhideColumn(1, 50);

        // Prepare PDF save options (default options are sufficient for this task)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Export the workbook to PDF
        workbook.Save("output.pdf", pdfOptions);
    }
}
